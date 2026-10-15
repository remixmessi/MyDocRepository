import {createHash,randomUUID} from 'node:crypto';
import {closeSync,existsSync,fsyncSync,lstatSync,mkdirSync,openSync,readFileSync,renameSync,unlinkSync,writeFileSync} from 'node:fs';
import {dirname,join} from 'node:path';
import type {SQLiteConnection} from './database';

type FileRecord={path:string;size:number;sha256:string};
export const attachmentDirectory=(databasePath:string)=>join(dirname(databasePath),'attachments');
const digest=(data:Uint8Array)=>createHash('sha256').update(data).digest('hex');
function filePath(directory:string,record:FileRecord){
 if(!/^[a-f0-9]{64}$/.test(record.sha256)||record.path!==record.sha256.slice(0,2)+'/'+record.sha256||!Number.isSafeInteger(record.size)||record.size<0||record.size>20*1024*1024)throw new Error('附件文件记录无效');
 for(const part of [directory,join(directory,record.sha256.slice(0,2))])if(existsSync(part)&&lstatSync(part).isSymbolicLink())throw new Error('附件目录不能是符号链接');
 return join(directory,record.path);
}
export function readAttachmentFile(directory:string,record:FileRecord){
 const path=filePath(directory,record);if(!lstatSync(path).isFile()||lstatSync(path).isSymbolicLink())throw new Error('附件文件类型无效');
 const bytes=readFileSync(path);if(bytes.length!==record.size||digest(bytes)!==record.sha256)throw new Error('附件文件校验失败：'+record.path);return bytes;
}
function writeAttachmentFile(directory:string,data:Uint8Array):FileRecord{
 if(data.length>20*1024*1024)throw new Error('附件超过 20 MB');
 const sha256=digest(data),record={path:sha256.slice(0,2)+'/'+sha256,size:data.length,sha256},path=filePath(directory,record);
 mkdirSync(dirname(path),{recursive:true,mode:0o700});
 if(existsSync(path)){readAttachmentFile(directory,record);return record;}
 const temp=path+'.'+randomUUID()+'.partial';
 try{const fd=openSync(temp,'wx',0o600);try{writeFileSync(fd,data);fsyncSync(fd);}finally{closeSync(fd);}renameSync(temp,path);if(process.platform!=='win32'){const dir=openSync(dirname(path),'r');try{fsyncSync(dir);}finally{closeSync(dir);}for(const parent of [directory,dirname(directory)]){const fd=openSync(parent,'r');try{fsyncSync(fd);}finally{closeSync(fd);}}}readAttachmentFile(directory,record);return record;}finally{if(existsSync(temp))unlinkSync(temp);}
}
/** Files are immutable and durable before committing their SQLite references. Rollback
 * may leave an unreferenced file, but never removes bytes needed by an older snapshot. */
export class AttachmentFiles {
 readonly directory:string;
 constructor(private db:SQLiteConnection,databasePath:string){this.directory=attachmentDirectory(databasePath);}
 get(scope:string,workspace:string,id:string){
  const row=this.db.prepare('SELECT name,mime,data FROM assets WHERE scope=? AND workspace=? AND id=?').get(scope,workspace,id) as {name:string;mime:string;data:Uint8Array}|undefined;
  if(!row)return undefined;
  const file=this.db.prepare('SELECT path,size,sha256 FROM asset_files WHERE scope=? AND workspace=? AND id=?').get(scope,workspace,id) as FileRecord|undefined;
  return {...row,data:file?readAttachmentFile(this.directory,file):Buffer.from(row.data)};
 }
 put(scope:string,workspace:string,id:string,name:string,mime:string,data:Uint8Array){
  const record=writeAttachmentFile(this.directory,data);
  this.db.transaction(()=>{
   this.db.prepare("INSERT INTO assets(scope,workspace,id,name,mime,data) VALUES (?,?,?,?,?,X'') ON CONFLICT(scope,workspace,id) DO UPDATE SET name=excluded.name,mime=excluded.mime,data=excluded.data").run(scope,workspace,id,name,mime);
   this.db.prepare('INSERT INTO asset_files VALUES (?,?,?,?,?,?) ON CONFLICT(scope,workspace,id) DO UPDATE SET path=excluded.path,size=excluded.size,sha256=excluded.sha256').run(scope,workspace,id,record.path,record.size,record.sha256);
  }).immediate();
 }
 remove(scope:string,workspace?:string,id?:string){
  const values:unknown[]=[scope];let where='scope=?';
  if(workspace!==undefined){where+=' AND workspace=?';values.push(workspace);}
  if(id!==undefined){if(workspace===undefined)throw new Error('附件工作区缺失');where+=' AND id=?';values.push(id);}
  this.db.transaction(()=>{this.db.prepare('DELETE FROM assets WHERE '+where).run(...values);this.db.prepare('DELETE FROM asset_files WHERE '+where).run(...values);}).immediate();
 }
 /** Called within the startup/restore write transaction, after schema checks. */
 migrate(){
  mkdirSync(this.directory,{recursive:true,mode:0o700});
  this.db.exec('DELETE FROM asset_files WHERE NOT EXISTS (SELECT 1 FROM assets a WHERE a.scope=asset_files.scope AND a.workspace=asset_files.workspace AND a.id=asset_files.id)');
  const rows=this.db.prepare('SELECT a.scope,a.workspace,a.id,a.name,a.mime FROM assets a LEFT JOIN asset_files f ON a.scope=f.scope AND a.workspace=f.workspace AND a.id=f.id WHERE f.id IS NULL').all() as {scope:string;workspace:string;id:string;name:string;mime:string}[];
  for(const row of rows){const asset=this.get(row.scope,row.workspace,row.id)!;this.put(row.scope,row.workspace,row.id,row.name,row.mime,asset.data);}
  for(const record of this.db.prepare('SELECT path,size,sha256 FROM asset_files').all() as FileRecord[])readAttachmentFile(this.directory,record);
 }
}
/** Make a snapshot self contained: preserve the established .sqlite + .json backup format. */
export function embedAttachmentFiles(db:SQLiteConnection,directory:string,schema:'main'|'attachment_snapshot'='main'){
 if(!db.prepare(`SELECT 1 FROM ${schema}.sqlite_master WHERE type='table' AND name='asset_files'`).get())return;
 db.transaction(()=>{
  for(const row of db.prepare(`SELECT scope,workspace,id,path,size,sha256 FROM ${schema}.asset_files f WHERE EXISTS (SELECT 1 FROM ${schema}.assets a WHERE a.scope=f.scope AND a.workspace=f.workspace AND a.id=f.id)`).all() as (FileRecord&{scope:string;workspace:string;id:string})[]){
   const data=readAttachmentFile(directory,row);db.prepare(`UPDATE ${schema}.assets SET data=? WHERE scope=? AND workspace=? AND id=?`).run(data,row.scope,row.workspace,row.id);
  }
  db.exec(`DELETE FROM ${schema}.asset_files`);
 }).immediate();
}
