import {it,expect} from 'vitest';
import Database from 'better-sqlite3';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {SCHEMA_MIGRATIONS,SCHEMA_VERSION} from '../packages/common/nbstore/src/schema';
import {upgradeCloudDatabase,CLOUD_SCHEMA_VERSION} from '../packages/frontend/apps/web/server/schema';

function fixture(){const dir=mkdtempSync(join(tmpdir(),'mydoc-startup-'));return {path:join(dir,'db.sqlite'),close:()=>rmSync(dir,{recursive:true,force:true})};}
const open=(path:string,cloud=false)=>new SQLiteProjectDatabase(path,path=>new Database(path),cloud?upgradeCloudDatabase:undefined);


it('旧库启动补建本地账号表，最新版本仍检查并修复缺表且保留数据',()=>{
 const f=fixture();try{
  const raw=new Database(f.path);raw.exec(SCHEMA_MIGRATIONS[1]+SCHEMA_MIGRATIONS[2]);raw.pragma('user_version=2');raw.prepare('INSERT INTO preferences VALUES (?,?,?)').run('main','theme','dark');raw.close();
  let db=open(f.path);expect(db.initializedThisStart).toBe(false);expect(db.db.pragma('user_version',{simple:true})).toBe(SCHEMA_VERSION);
  expect(db.db.prepare('SELECT count(*) n FROM local_accounts').get()).toEqual({n:0});
  db.db.exec('DROP TABLE drafts');db.close();db=open(f.path);
  expect(db.db.prepare('SELECT count(*) n FROM drafts').get()).toEqual({n:0});expect(db.db.prepare('SELECT value FROM preferences').get()).toEqual({value:'dark'});db.close();
 }finally{f.close();}
});

it('最新库缺少基础字段时拒绝启动，补表与版本变更整体回滚',()=>{
 const f=fixture();try{
  const db=open(f.path);db.db.exec('DROP TABLE drafts; ALTER TABLE local_accounts DROP COLUMN name');db.close();
  expect(()=>open(f.path)).toThrow('local_accounts 缺少字段 name');
  const raw=new Database(f.path);expect(raw.pragma('user_version',{simple:true})).toBe(SCHEMA_VERSION);expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='drafts'").get()).toBeUndefined();raw.close();
 }finally{f.close();}
});

it('协作旧版本增量补字段、会话 ID 和迁移记录，重启不覆盖既有数据',()=>{
 const f=fixture();try{
  let db=open(f.path,true);db.db.prepare('INSERT INTO cloud_sessions(token,user,expires) VALUES (?,?,?)').run('token','user',100);
  db.db.exec('DELETE FROM cloud_migrations WHERE version>4; ALTER TABLE cloud_releases DROP COLUMN arch; ALTER TABLE cloud_users DROP COLUMN avatar; ALTER TABLE cloud_sessions DROP COLUMN agent');db.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:CLOUD_SCHEMA_VERSION});
  const session=db.db.prepare('SELECT id,user,agent FROM cloud_sessions').get() as {id:string;user:string;agent:null};expect(session.user).toBe('user');expect(session.id).toHaveLength(36);expect(session.agent).toBeNull();
  expect(db.db.prepare('SELECT avatar FROM cloud_users LIMIT 0').all()).toEqual([]);expect(db.db.prepare('SELECT arch FROM cloud_releases LIMIT 0').all()).toEqual([]);db.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT id FROM cloud_sessions').get()).toEqual({id:session.id});db.close();
 }finally{f.close();}
});

it('高版本协作库拒绝启动，共用表迁移也不会留下修改',()=>{
 const f=fixture();try{
  const raw=new Database(f.path);raw.exec(SCHEMA_MIGRATIONS[1]+SCHEMA_MIGRATIONS[2]+'CREATE TABLE cloud_migrations(version INTEGER PRIMARY KEY);INSERT INTO cloud_migrations VALUES (99)');raw.pragma('user_version=2');raw.close();
  expect(()=>open(f.path,true)).toThrow('协作数据库版本高于当前程序');
  const check=new Database(f.path);expect(check.pragma('user_version',{simple:true})).toBe(2);expect(check.prepare("SELECT name FROM sqlite_master WHERE name IN ('local_accounts','cloud_users')").all()).toEqual([]);check.close();
 }finally{f.close();}
});


it('办公恢复回执从版本 8 原子升级，重启与快照保留，失败回滚',()=>{
 const f=fixture();try{
  let db=open(f.path,true);db.db.prepare('INSERT INTO preferences VALUES (?,?,?)').run('main','keep','old');db.db.exec('DROP TABLE cloud_office_restores;DROP TABLE cloud_workspace_restores;DROP TABLE cloud_office_versions;DROP TABLE cloud_office_exports;DELETE FROM cloud_migrations WHERE version>=9');db.close();
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),connection=>{upgradeCloudDatabase(connection);throw new Error('receipt migration failure');})).toThrow('receipt migration failure');
  const raw=new Database(f.path);expect(raw.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:8});expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='cloud_office_restores'").get()).toBeUndefined();raw.close();
  db=open(f.path,true);db.db.prepare('INSERT INTO cloud_office_restores VALUES (?,?,?,?,?,?)').run('space','actor','operation','hash','resource',1);expect(db.db.prepare('SELECT value FROM preferences').get()).toEqual({value:'old'});db.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT resource FROM cloud_office_restores').get()).toEqual({resource:'resource'});const snapshot=join(f.path+'.snapshot');db.db.exec("VACUUM INTO '"+snapshot+"'");db.close();
  const restored=open(snapshot,true);expect(restored.db.prepare('SELECT operation FROM cloud_office_restores').get()).toEqual({operation:'operation'});restored.close();
 }finally{f.close();}
});


it('完整工作区恢复回执从版本 9 原子升级、保留数据、重启及物理备份兼容',()=>{
 const f=fixture();try{
  let db=open(f.path,true);db.db.prepare('INSERT INTO preferences VALUES (?,?,?)').run('main','kept','yes');db.db.exec('DROP TABLE cloud_workspace_restores;DROP TABLE cloud_office_versions;DELETE FROM cloud_migrations WHERE version>=10');db.close();
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),connection=>{upgradeCloudDatabase(connection);throw new Error('archive migration failed');})).toThrow('archive migration failed');
  const raw=new Database(f.path);expect(raw.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:9});expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='cloud_workspace_restores'").get()).toBeUndefined();raw.close();
  db=open(f.path,true);db.db.prepare('INSERT INTO cloud_workspace_restores VALUES (?,?,?,?,?)').run('actor','operation','hash','space',1);db.close();db=open(f.path,true);
  expect(db.db.prepare('SELECT value FROM preferences').get()).toEqual({value:'yes'});expect(db.db.prepare('SELECT count(*) n FROM cloud_migrations WHERE version=10').get()).toEqual({n:1});
  const snapshot=f.path+'.backup';db.db.exec("VACUUM INTO '"+snapshot+"'");db.close();const restored=open(snapshot,true);expect(restored.db.prepare('SELECT operation FROM cloud_workspace_restores').get()).toEqual({operation:'operation'});restored.close();
 }finally{f.close();}
});


it('命名版本从协作版本 10 原子升级，历史记录和物理备份保留，重复启动不重复迁移',()=>{
 const f=fixture();try{
  let db=open(f.path,true);db.db.prepare('INSERT INTO preferences VALUES (?,?,?)').run('main','version-keep','yes');db.db.exec('DROP TABLE cloud_office_versions;DROP TABLE cloud_office_exports;DELETE FROM cloud_migrations WHERE version>=11');db.close();
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),connection=>{upgradeCloudDatabase(connection);throw new Error('version migration failed');})).toThrow('version migration failed');const raw=new Database(f.path);expect(raw.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:10});expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='cloud_office_versions'").get()).toBeUndefined();raw.close();
  db=open(f.path,true);db.db.prepare('INSERT INTO cloud_office_versions VALUES (?,?,?,?,?,?,?,?,?,?)').run('space','doc:source','version','named','actor','operation',1,Buffer.from([1,2,3]),'hash',0);db.close();db=open(f.path,true);expect(db.db.prepare('SELECT value FROM preferences').get()).toEqual({value:'yes'});expect(db.db.prepare('SELECT count(*) n FROM cloud_migrations WHERE version=11').get()).toEqual({n:1});const snapshot=f.path+'.versions';db.db.exec("VACUUM INTO '"+snapshot+"'");db.close();const restored=open(snapshot,true);expect(restored.db.prepare('SELECT name,data FROM cloud_office_versions').get()).toEqual({name:'named',data:Buffer.from([1,2,3])});restored.close();
 }finally{f.close();}
});

it('导出任务表从协作版本 11 原子升级：保留既有版本与资源、重复启动幂等、失败回滚、拒绝更高版本',()=>{
 const f=fixture();try{
  let db=open(f.path,true);db.db.prepare("INSERT INTO cloud_office_versions(workspace,resource,id,name,user,operation,created,data,sha256) VALUES ('s','r','v','原稿','u','op',1,x'00','h')").run();db.db.exec('DROP TABLE cloud_office_exports;DROP TABLE cloud_notifications;DROP TABLE cloud_office_activity;DELETE FROM cloud_migrations WHERE version>=12');db.close();
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),connection=>{upgradeCloudDatabase(connection);throw new Error('export migration failure');})).toThrow('export migration failure');
  const raw=new Database(f.path);expect(raw.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:11});expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='cloud_office_exports'").get()).toBeUndefined();raw.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:CLOUD_SCHEMA_VERSION});expect(db.db.prepare('SELECT count(*) n FROM cloud_migrations WHERE version=12').get()).toEqual({n:1});expect(db.db.prepare('SELECT name FROM cloud_office_versions').get()).toEqual({name:'原稿'});
  db.db.prepare("INSERT INTO cloud_office_exports(workspace,resource,id,user,operation,format,scope,status,created,snapshot,filename,mime) VALUES ('s','r','e','u','o','docx','paragraphs','running',1,'h','a.docx','m')").run();db.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT count(*) n FROM cloud_migrations WHERE version=12').get()).toEqual({n:1});expect(db.db.prepare('SELECT status FROM cloud_office_exports').get()).toEqual({status:'running'});db.db.prepare('INSERT INTO cloud_migrations VALUES (?)').run(CLOUD_SCHEMA_VERSION+1);db.close();
  expect(()=>open(f.path,true)).toThrow('高于当前程序');
 }finally{f.close();}
});

it('协作通知与编辑记录表从协作版本 12 原子升级：保留既有数据、重复启动幂等、失败回滚、拒绝更高版本',()=>{
 const f=fixture();try{
  let db=open(f.path,true);db.db.prepare("INSERT INTO cloud_office_exports(workspace,resource,id,user,operation,format,scope,status,created,snapshot,filename,mime) VALUES ('s','r','e','u','o','docx','paragraphs','done',1,'h','a.docx','m')").run();db.db.prepare("INSERT INTO cloud_workspace_audit(id,workspace,actor,actor_name,action,details,created) VALUES ('a','s','u','用户','rename','保留',1)").run();
  db.db.exec('DROP TABLE cloud_notifications;DROP TABLE cloud_office_activity;DELETE FROM cloud_migrations WHERE version=13');db.close();
  expect(()=>new SQLiteProjectDatabase(f.path,path=>new Database(path),connection=>{upgradeCloudDatabase(connection);throw new Error('collaboration migration failure');})).toThrow('collaboration migration failure');
  const raw=new Database(f.path);expect(raw.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:12});expect(raw.prepare("SELECT name FROM sqlite_master WHERE name IN ('cloud_notifications','cloud_office_activity')").all()).toEqual([]);raw.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:13});expect(db.db.prepare('SELECT status FROM cloud_office_exports').get()).toEqual({status:'done'});expect(db.db.prepare('SELECT details FROM cloud_workspace_audit').get()).toEqual({details:'保留'});
  db.db.prepare("INSERT INTO cloud_notifications(id,user,workspace,resource,kind,actor,actor_name,thread,excerpt,created) VALUES ('n','u','s','r','mention','a','甲','t','摘要',1)").run();db.db.prepare("INSERT INTO cloud_office_activity(workspace,resource,user,user_name,started,last) VALUES ('s','r','u','甲',1,2)").run();
  expect(()=>db.db.prepare("INSERT INTO cloud_notifications(id,user,workspace,resource,kind,actor,actor_name,thread,excerpt,created) VALUES ('x','u','s','r','other','a','甲','t','',1)").run()).toThrow();db.close();
  db=open(f.path,true);expect(db.db.prepare('SELECT count(*) n FROM cloud_migrations WHERE version=13').get()).toEqual({n:1});expect(db.db.prepare('SELECT excerpt,read FROM cloud_notifications').get()).toEqual({excerpt:'摘要',read:0});expect(db.db.prepare('SELECT edits FROM cloud_office_activity').get()).toEqual({edits:1});
  db.db.exec('DROP TABLE cloud_office_activity');db.close();db=open(f.path,true);expect(db.db.prepare('SELECT count(*) n FROM cloud_office_activity').get()).toEqual({n:0});expect(db.db.prepare('SELECT count(*) n FROM cloud_notifications').get()).toEqual({n:1});
  db.db.prepare('INSERT INTO cloud_migrations VALUES (?)').run(CLOUD_SCHEMA_VERSION+1);db.close();expect(()=>open(f.path,true)).toThrow('高于当前程序');
 }finally{f.close();}
});
