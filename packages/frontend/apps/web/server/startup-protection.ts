import {attachmentDirectory,embedAttachmentFiles} from '../../../../common/nbstore/server/attachment-files';
import {createHash,randomUUID} from 'node:crypto';
import {chmodSync,closeSync,existsSync,fsyncSync,mkdirSync,openSync,renameSync,rmSync,writeFileSync} from 'node:fs';
import {dirname,join} from 'node:path';
import type {SQLiteConnection} from '../../../../common/nbstore/server/database';
import {CLOUD_SCHEMA_VERSION} from './schema';
import {assertPersistentStorage} from './persistent-storage';
import {protectIncomingWal} from './startup-files';

/** Container upgrades never create a replacement for an existing database. */
export function hostedDatabaseProtection(path:string,env:NodeJS.ProcessEnv=process.env){
 assertPersistentStorage(path,env);
 if(env.MYDOC_CONTAINER!=='1')return {};
 const existing=existsSync(path);
 if(existing)protectIncomingWal(path);
 return {mustExist:existing,beforeMigration:(db:SQLiteConnection)=>{if(existing)snapshotBeforeCloudMigration(path,db,env);}};
}

export function snapshotBeforeCloudMigration(path:string,db:SQLiteConnection,env:NodeJS.ProcessEnv){
  const tables=new Set((db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all() as {name:string}[]).map(row=>row.name));
  for(const table of ['documents','assets','preferences','cloud_users','cloud_workspaces','cloud_members','cloud_migrations']){
   if(!tables.has(table))throw new Error(`已有数据库缺少 ${table}，拒绝补建空表。请恢复包含 WAL 的完整数据库备份。`);
  }
  const version=(db.prepare('SELECT max(version) version FROM cloud_migrations').get() as {version:number|null}).version??0;
  if(version>CLOUD_SCHEMA_VERSION)throw new Error('协作数据库版本高于当前程序，请升级程序后再启动');
  const check=db.pragma('quick_check') as {quick_check:string}[];
  if(check.length!==1||check[0].quick_check!=='ok')throw new Error('数据库完整性检查失败，停止启动');
  const generation=(db.prepare("SELECT value FROM preferences WHERE scope='__system' AND key='restore-generation'").get() as {value:string}|undefined)?.value??'';
  const counts=Object.fromEntries(['cloud_users','cloud_workspaces','cloud_members','documents','assets'].map(table=>[table,(db.prepare(`SELECT count(*) n FROM ${table}`).get() as {n:number}).n]));
  if(!counts.cloud_users&&!counts.cloud_workspaces&&!counts.documents&&!counts.assets&&env.MYDOC_INIT_DATABASE!=='1')throw new Error('已有数据库没有账号、工作空间或文档，拒绝将空库当成升级成功。请检查恢复文件和挂载目录。');
  const imageVersion=env.MYDOC_VERSION??'unversioned',schema=db.pragma('user_version',{simple:true});
  const identities=db.prepare('SELECT id FROM cloud_users UNION SELECT id FROM cloud_workspaces ORDER BY id').all();
  const key=createHash('sha256').update(JSON.stringify({imageVersion,schema,version,generation,identities})).digest('hex').slice(0,24);
  const directory=join(dirname(path),'upgrade-backups'),snapshot=join(directory,`before-upgrade-${key}.sqlite`);
  if(existsSync(snapshot)&&existsSync(snapshot+'.json'))return;
  mkdirSync(directory,{recursive:true,mode:0o700});
  const temporary=join(directory,randomUUID()+'.partial');
  try{
   // VACUUM INTO takes a consistent SQLite snapshot including committed WAL.
   // Copying cloud.sqlite alone can silently omit the latest committed pages.
   db.prepare('VACUUM INTO ?').run(temporary);
   db.prepare('ATTACH DATABASE ? AS attachment_snapshot').run(temporary);
   try{embedAttachmentFiles(db,attachmentDirectory(path),'attachment_snapshot');}finally{db.exec('DETACH DATABASE attachment_snapshot');}
   chmodSync(temporary,0o600);
   const descriptor=openSync(temporary,'r+');try{fsyncSync(descriptor);}finally{closeSync(descriptor);}
   renameSync(temporary,snapshot);
   writeFileSync(temporary,JSON.stringify({createdAt:new Date().toISOString(),imageVersion,schemaVersion:schema,cloudSchemaVersion:version,counts},null,2),{mode:0o600,flag:'wx'});
   renameSync(temporary,snapshot+'.json');
   console.log('MyDoc upgrade safety snapshot: '+snapshot+' '+JSON.stringify(counts));
  }finally{rmSync(temporary,{force:true});}
}

