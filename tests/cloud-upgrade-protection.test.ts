import {it,expect} from 'vitest';
import Database from 'better-sqlite3';
import {existsSync,mkdtempSync,readdirSync,rmSync,writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import * as Y from 'yjs';
import {SQLiteProjectDatabase,type SQLiteConnection} from '../packages/common/nbstore/server/database';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {upgradeCloudDatabase} from '../packages/frontend/apps/web/server/schema';
import {snapshotBeforeCloudMigration} from '../packages/frontend/apps/web/server/startup-protection';
import {serverConfig} from '../packages/frontend/apps/web/server/config';

const admin={email:'original@example.com',password:'original-password-123',name:'原管理员'};
function fixture(){const dir=mkdtempSync(join(tmpdir(),'mydoc-cloud-upgrade-')),path=join(dir,'cloud.sqlite');const cloud=new CloudServer(path,{defaultAdmin:admin});const doc=new Y.Doc();doc.getMap('body').set('text','WAL 内的升级前正文');cloud.database.append('user:original','doc:original',Y.encodeStateAsUpdate(doc));doc.destroy();cloud.database.db.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?)').run('user:original','local','original','附件.txt','text/plain',Buffer.from('原附件字节'));cloud.database.db.prepare('INSERT INTO cloud_workspaces VALUES (?,?,?)').run('original-workspace','原团队',1);cloud.database.db.prepare('INSERT INTO cloud_members VALUES (?,?,?)').run('original-workspace','original-user','owner');cloud.database.db.exec('DROP TABLE local_account_sessions; DROP TABLE local_accounts; PRAGMA user_version=2');return {dir,path,cloud,close(){cloud.close();rmSync(dir,{recursive:true,force:true});}};}
const protection=(path:string)=>({beforeMigration:(db:SQLiteConnection)=>snapshotBeforeCloudMigration(path,db,{MYDOC_VERSION:'test-version'})});

it('升级前快照包含尚在 WAL 的正文及附件，保留旧结构供旧镜像回滚；重复启动不重复备份',()=>{
 const f=fixture();try{
  const open=()=>new SQLiteProjectDatabase(f.path,path=>new Database(path,{fileMustExist:true}),upgradeCloudDatabase,protection(f.path));
  let db=open();expect(db.schemaVersion).toBe(3);db.close();
  let files=readdirSync(join(f.dir,'upgrade-backups')).filter(file=>file.endsWith('.sqlite'));expect(files).toHaveLength(1);
  const snapshot=new Database(join(f.dir,'upgrade-backups',files[0]),{readonly:true,fileMustExist:true});
  expect(snapshot.pragma('user_version',{simple:true})).toBe(2);expect(snapshot.pragma('quick_check',{simple:true})).toBe('ok');
  expect(snapshot.prepare('SELECT data FROM documents').get()).toEqual(f.cloud.database.db.prepare('SELECT data FROM documents').get());
  expect(snapshot.prepare('SELECT name,data FROM assets').get()).toEqual({name:'附件.txt',data:Buffer.from('原附件字节')});
  expect(snapshot.prepare('SELECT email,name FROM cloud_users').get()).toEqual({email:admin.email,name:admin.name});snapshot.close();
  db=open();db.close();const afterSecond=readdirSync(join(f.dir,'upgrade-backups')).length;db=open();db.close();expect(readdirSync(join(f.dir,'upgrade-backups'))).toHaveLength(afterSecond);
 }finally{f.close();}
});

it('破坏原文档、附件、账号和空间的迁移回滚，原值及旧版本号保留',()=>{
 for(const table of ['documents','assets','cloud_users','cloud_workspaces','cloud_members']){
  const f=fixture();try{
   const before=f.cloud.database.db.prepare(`SELECT * FROM ${table}`).all();
   expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),db=>{upgradeCloudDatabase(db);db.exec(`DELETE FROM ${table}`);},{preserve:{cloud_users:'id,email,name,salt,password',cloud_workspaces:'id,name,created',cloud_members:'workspace,user,role'}})).toThrow('已回滚');
   expect(f.cloud.database.db.prepare(`SELECT * FROM ${table}`).all()).toEqual(before);expect(f.cloud.database.db.pragma('user_version',{simple:true})).toBe(2);
  }finally{f.close();}
 }
});

it('保留文档主键和条数却清空二进制正文也会被拦截，不能用空内容覆盖旧文档',()=>{
 const f=fixture();try{
  const before=f.cloud.database.db.prepare('SELECT * FROM documents').all();
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),db=>db.exec("UPDATE documents SET data=X'0000'"))).toThrow('已回滚');
  expect(f.cloud.database.db.prepare('SELECT * FROM documents').all()).toEqual(before);
  expect(f.cloud.database.db.pragma('user_version',{simple:true})).toBe(2);
 }finally{f.close();}
});

it('已有空库拒绝升级；文件在挂载检查之后消失时也不自动重建',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-empty-upgrade-')),path=join(dir,'cloud.sqlite');
 try{
  const empty=new CloudServer(path);empty.close();
  expect(()=>new SQLiteProjectDatabase(path,path=>new Database(path),upgradeCloudDatabase,protection(path))).toThrow('拒绝将空库当成升级成功');
  rmSync(path);expect(()=>new CloudServer(path,{mustExist:true})).toThrow();expect(existsSync(path)).toBe(false);
 }finally{rmSync(dir,{recursive:true,force:true});}
});

it('快照写入失败、基础表缺失和协作版本过高均停止启动，不创建空表或推进版本',()=>{
 const f=fixture();try{
  writeFileSync(join(f.dir,'upgrade-backups'),'blocked');
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),upgradeCloudDatabase,protection(f.path))).toThrow();
  expect(f.cloud.database.db.pragma('user_version',{simple:true})).toBe(2);
  rmSync(join(f.dir,'upgrade-backups'));f.cloud.database.db.exec('DROP TABLE documents');
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),upgradeCloudDatabase,protection(f.path))).toThrow('拒绝补建空表');
  expect(f.cloud.database.db.prepare("SELECT name FROM sqlite_master WHERE name='documents'").get()).toBeUndefined();
 }finally{f.close();}
 const high=fixture();try{
  high.cloud.database.db.exec('INSERT INTO cloud_migrations VALUES(999)');
  expect(()=>new SQLiteProjectDatabase(high.path,path=>new Database(path),upgradeCloudDatabase,protection(high.path))).toThrow('协作数据库版本高于');
  expect(high.cloud.database.db.pragma('user_version',{simple:true})).toBe(2);
 }finally{high.close();}
});

it('容器数据库路径环境变量覆盖旧配置文件，不会随工作目录切换到另一个库',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-path-')),path=join(dir,'server-config.json');
 try{
  writeFileSync(path,JSON.stringify({cloudDbPath:'data/old.sqlite',host:'127.0.0.1',port:9090}));
  expect(serverConfig({MYDOC_CONFIG_PATH:path,MYDOC_CLOUD_DB_PATH:'/app/data/cloud.sqlite',MYDOC_HOST:'0.0.0.0',MYDOC_PORT:'8080'})).toMatchObject({cloudDbPath:'/app/data/cloud.sqlite',host:'0.0.0.0',port:8080});
 }finally{rmSync(dir,{recursive:true,force:true});}
});
