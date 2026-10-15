import {it,expect} from 'vitest';
import Database from 'better-sqlite3';
import {copyFileSync,existsSync,mkdirSync,mkdtempSync,readdirSync,readFileSync,rmSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {protectIncomingWal} from '../packages/frontend/apps/web/server/startup-files';
import {snapshotBeforeCloudMigration} from '../packages/frontend/apps/web/server/startup-protection';

function fixture(){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-input-wal-')),path=join(dir,'cloud.sqlite'),writerPath=join(dir,'writer.sqlite');
 const original=new CloudServer(path,{defaultAdmin:{email:'original@example.com',password:'original-password-123',name:'原账号'}});
 const user=(original.database.db.prepare('SELECT id FROM cloud_users').get() as {id:string}).id;
 original.database.db.prepare('INSERT INTO cloud_workspaces VALUES(?,?,?)').run('original-space','原团队',1);original.close();
 copyFileSync(path,writerPath);const writer=new Database(writerPath);writer.pragma('wal_autocheckpoint=0');
 return {dir,path,user,writer,attachWal(){copyFileSync(writerPath+'-wal',path+'-wal');},close(){writer.close();rmSync(dir,{recursive:true,force:true});}};
}
function bytes(path:string){return Object.fromEntries(['','-wal','-shm','-journal'].filter(suffix=>existsSync(path+suffix)).map(suffix=>[suffix,readFileSync(path+suffix)]));}

it('复现恢复主库混入另一份 WAL：完整性检查仍为 ok，但原账号和空间一起消失；新检查在打开原库前拦截',()=>{
 const f=fixture();try{
  f.writer.transaction(()=>{
   f.writer.exec('DELETE FROM cloud_users;DELETE FROM cloud_workspaces;DELETE FROM cloud_bootstrap');
   f.writer.prepare('INSERT INTO cloud_users(id,email,name,salt,password) VALUES(?,?,?,?,?)').run('replacement-admin','replacement@example.com','新初始化管理员','salt','password');
  })();f.attachWal();const before=bytes(f.path);
  // Reproduce the prior startup behavior on a disposable copy only. It passes
  // quick_check and the nonempty-database check because a replacement admin exists.
  const simulation=join(f.dir,'simulation');mkdirSync(simulation);const copy=join(simulation,'cloud.sqlite');copyFileSync(f.path,copy);copyFileSync(f.path+'-wal',copy+'-wal');
  const previous=new CloudServer(copy,{beforeMigration:db=>snapshotBeforeCloudMigration(copy,db,{MYDOC_VERSION:'previous'})});
  expect(previous.database.db.pragma('quick_check',{simple:true})).toBe('ok');
  expect(previous.database.db.prepare('SELECT id FROM cloud_users').all()).toEqual([{id:'replacement-admin'}]);
  expect(previous.database.db.prepare('SELECT count(*) n FROM cloud_workspaces').get()).toEqual({n:0});previous.close();
  expect(()=>protectIncomingWal(f.path)).toThrow('WAL 应用后已有账号消失');
  expect(bytes(f.path)).toEqual(before);
  const saved=join(f.dir,'upgrade-inputs',readdirSync(join(f.dir,'upgrade-inputs')).find(name=>name.startsWith('input-'))!);
  expect(readFileSync(join(saved,'database.sqlite'))).toEqual(before['']);expect(readFileSync(join(saved,'database.sqlite-wal'))).toEqual(before['-wal']);
 }finally{f.close();}
});

it('同一库正常崩溃恢复保留新注册账号、改名改邮箱及删除空间，不修改原始主库和 WAL',()=>{
 const f=fixture();try{
  f.writer.prepare('UPDATE cloud_users SET name=?,email=? WHERE id=?').run('改名','changed@example.com',f.user);
  f.writer.prepare('INSERT INTO cloud_users(id,email,name,salt,password) VALUES(?,?,?,?,?)').run('new-user','new@example.com','新用户','salt','password');
  f.writer.exec('DELETE FROM cloud_workspaces');f.attachWal();const before=bytes(f.path);
  expect(()=>protectIncomingWal(f.path)).not.toThrow();expect(bytes(f.path)).toEqual(before);
  const first=readdirSync(join(f.dir,'upgrade-inputs'));expect(()=>protectIncomingWal(f.path)).not.toThrow();expect(readdirSync(join(f.dir,'upgrade-inputs'))).toEqual(first);
 }finally{f.close();}
});

it('首次建表和账号只在 WAL 中时允许正常恢复；没有 WAL 时不创建原始文件归档目录',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-first-wal-')),writerPath=join(dir,'writer.sqlite'),path=join(dir,'cloud.sqlite');
 const writer=new Database(writerPath);try{
  writer.pragma('journal_mode=WAL');writer.pragma('wal_autocheckpoint=0');
  writer.exec("CREATE TABLE cloud_users(id TEXT PRIMARY KEY);INSERT INTO cloud_users VALUES('first-user')");
  copyFileSync(writerPath,path);expect(()=>protectIncomingWal(path)).not.toThrow();expect(existsSync(join(dir,'upgrade-inputs'))).toBe(false);
  copyFileSync(writerPath+'-wal',path+'-wal');const before=bytes(path);
  expect(()=>protectIncomingWal(path)).not.toThrow();expect(bytes(path)).toEqual(before);
 }finally{writer.close();rmSync(dir,{recursive:true,force:true});}
});
