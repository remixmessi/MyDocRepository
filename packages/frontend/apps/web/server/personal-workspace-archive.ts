import * as Y from 'yjs';
import {createHash} from 'node:crypto';
import {restoreLocalOfficeArchive,WorkspaceArchiveError} from './workspace-archive';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import type {CloudWorkspaceArchive} from '../../../../common/nbstore/src/workspace-archive';

function registry(database:SQLiteProjectDatabase,scope:string){const doc=new Y.Doc(),bytes=database.load(scope,'registry:local');if(bytes)Y.applyUpdate(doc,bytes);return doc;}
export function personalRestoreReceipt(database:SQLiteProjectDatabase,scope:string,operation:string,sha256:string){
 const row=database.db.prepare('SELECT value FROM preferences WHERE scope=? AND key=?').get(scope,'personal-workspace-restore:'+operation) as {value:string}|undefined;if(!row)return;
 const receipt=JSON.parse(row.value) as {sha256:string;id:string};if(receipt.sha256!==sha256)throw new WorkspaceArchiveError('恢复编号已用于其他备份',409);const doc=registry(database,scope);try{if(!doc.getMap('workspaces').has(receipt.id))throw new WorkspaceArchiveError('此前恢复的工作区已删除，请选择备份开始新恢复',410);return {id:receipt.id,registryUpdate:Buffer.from(Y.encodeStateAsUpdate(doc)).toString('base64'),repeated:true};}finally{doc.destroy();}
}
/** Registry, note states/assets and the office namespace become visible in one SQLite transaction. */
export function restorePersonalWorkspaceArchive(database:SQLiteProjectDatabase,scope:string,user:string,operation:string,sha256:string,archive:CloudWorkspaceArchive){
 return database.db.transaction(()=>{
  const previous=personalRestoreReceipt(database,scope,operation,sha256);if(previous)return previous;const id=archive.workspace.id,doc=registry(database,scope);try{
   if(doc.getMap('workspaces').size>=10000)throw new WorkspaceArchiveError('个人工作区数量达到上限',413);
   if(doc.getMap('workspaces').has(id))throw new WorkspaceArchiveError('恢复目标已存在',409);
   const officeIds=new Set(archive.office.map(file=>file.id)),officeWorkspace='local-office-'+createHash('sha256').update(JSON.stringify([scope,id])).digest('hex').slice(0,64),officeAssets=archive.assets.filter(asset=>asset.doc!==null&&officeIds.has(asset.doc));
   restoreLocalOfficeArchive(database,officeWorkspace,user,{...archive,workspace:{...archive.workspace,id:officeWorkspace},documents:archive.documents.filter(entry=>officeIds.has(entry.id)),assets:officeAssets});
   for(const entry of archive.documents)if(!officeIds.has(entry.id)){if(entry.id!=='workspace:'+id&&database.load(scope,entry.id))throw new WorkspaceArchiveError('恢复文档标识冲突，请重试',409);database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run(scope,entry.id,Buffer.from(entry.state.data,'base64'));}
   for(const asset of archive.assets)if(!officeAssets.includes(asset))database.attachments.put(scope,id,asset.id,asset.name,asset.mime,Buffer.from(asset.state.data,'base64'));
   doc.getMap('workspaces').set(id,{id,name:archive.workspace.name.slice(0,156)+'（恢复）',createdAt:Date.now()});database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?) ON CONFLICT(scope,id) DO UPDATE SET data=excluded.data').run(scope,'registry:local',Buffer.from(Y.encodeStateAsUpdate(doc)));database.preference(scope,'personal-workspace-restore:'+operation,JSON.stringify({sha256,id}));return {id,registryUpdate:Buffer.from(Y.encodeStateAsUpdate(doc)).toString('base64'),repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}
