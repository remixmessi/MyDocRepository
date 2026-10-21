import {readOfficeBackup,encodeOfficeBackup} from './office-backup';
import {remapOfficeMedia} from './office-media';
import {randomUUID,createHash} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {MAX_OFFICE_VERSIONS,MAX_OFFICE_VERSION_BYTES} from '../../../../common/nbstore/src/office-version';
export class OfficeVersionError extends Error{constructor(message:string,readonly status=400){super(message);}}
export function findOfficeVersionOperation(database:SQLiteProjectDatabase,workspace:string,resource:string,user:string,operation:string,name:string){
 const row=database.db.prepare('SELECT id,name,deleted FROM cloud_office_versions WHERE workspace=? AND resource=? AND user=? AND operation=?').get(workspace,resource,user,operation) as {id:string;name:string;deleted:number}|undefined;
 if(row){if(row.name!==name)throw new OfficeVersionError('操作编号已用于另一版本名称',409);if(row.deleted)throw new OfficeVersionError('先前保存的版本已删除，请开始新的保存',410);}return row;
}
export function saveOfficeVersion(database:SQLiteProjectDatabase,workspace:string,resource:string,user:string,operation:string,name:string,data:Uint8Array){
 return database.db.transaction(()=>{
  const receipt=findOfficeVersionOperation(database,workspace,resource,user,operation,name);if(receipt)return {id:receipt.id,repeated:true};
  if(!database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,resource))throw new OfficeVersionError('源文件已删除',404);
  const count=database.db.prepare('SELECT count(*) count FROM cloud_office_versions WHERE workspace=? AND resource=? AND deleted=0').get(workspace,resource) as {count:number},stats=database.db.prepare('SELECT count(*) count,coalesce(sum(length(data)),0) bytes FROM cloud_office_versions WHERE workspace=?').get(workspace) as {count:number;bytes:number};
  if(count.count>=MAX_OFFICE_VERSIONS)throw new OfficeVersionError('单文件最多保留 100 个版本，请删除不需要的版本',413);if(stats.count>=10000)throw new OfficeVersionError('工作区版本操作记录已达上限，请备份恢复为新工作区继续',413);if(stats.bytes+data.length>MAX_OFFICE_VERSION_BYTES)throw new OfficeVersionError('工作区版本记录或 200 MB 容量已满，请先清理历史版本',413);
  const id=randomUUID();database.db.prepare('INSERT INTO cloud_office_versions(workspace,resource,id,name,user,operation,created,data,sha256,deleted) VALUES (?,?,?,?,?,?,?,?,?,0)').run(workspace,resource,id,name,user,operation,Date.now(),Buffer.from(data),createHash('sha256').update(data).digest('hex'));return {id,repeated:false};
 }).immediate();
}
export function officeVersionData(database:SQLiteProjectDatabase,workspace:string,resource:string,id:string){
 const row=database.db.prepare('SELECT name,data,sha256,deleted,user FROM cloud_office_versions WHERE workspace=? AND resource=? AND id=?').get(workspace,resource,id) as {name:string;data:Buffer|null;sha256:string;deleted:number;user:string}|undefined;
 if(!row)throw new OfficeVersionError('版本不存在',404);if(row.deleted||!row.data)throw new OfficeVersionError('版本已删除',410);if(createHash('sha256').update(row.data).digest('hex')!==row.sha256)throw new OfficeVersionError('版本校验失败，请从完整备份恢复',409);return {...row,data:Buffer.from(row.data)};
}

/** Atomic generation switch with an automatic safety snapshot and an idempotent receipt. */
export function revertOfficeVersion(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:{operationId:string;version:string;generation:string;base:string},parsed:import('./office-backup').OfficeBackup,state:Uint8Array){
 const signature=createHash('sha256').update(JSON.stringify(['revert',id,input.version,input.generation,input.base])).digest('hex');
 return database.db.transaction(()=>{
  const receipt=database.db.prepare('SELECT sha256,resource FROM cloud_office_restores WHERE workspace=? AND user=? AND operation=?').get(workspace,user,input.operationId) as {sha256:string;resource:string}|undefined;
  const resource=database.db.prepare('SELECT generation,kind,source_format AS sourceFormat,original_asset_id AS originalAssetId,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {generation:string;kind:import('../../../../common/nbstore/src/office').OfficeKind;sourceFormat:string;originalAssetId:string|null;deleted:number}|undefined;
  if(!resource||resource.deleted)throw new OfficeVersionError('文件不存在或已删除',404);
  if(receipt){if(receipt.sha256!==signature||receipt.resource!==id)throw new OfficeVersionError('操作编号已用于另一恢复操作',409);return {id,generation:resource.generation,repeated:true};}
  if(resource.generation!==input.generation)throw new OfficeVersionError('文件代际已更新，请重新载入并比较后恢复',409);
  const current=database.load('cloud:'+workspace,id);if(!current||createHash('sha256').update(current).digest('hex')!==input.base)throw new OfficeVersionError('当前文件已有新修改，请重新比较后恢复',409);
  const version=officeVersionData(database,workspace,id,input.version);const snapshot=readOfficeBackup(database,workspace,id);
  if(resource.kind!==parsed.resource.kind||resource.sourceFormat!==parsed.resource.sourceFormat||snapshot.archive.original?.sha256!==parsed.original?.sha256)throw new OfficeVersionError('版本原件或格式与当前资源不一致');
  // This snapshot and the generation change commit together; quota/IO failures leave the head untouched.
  saveOfficeVersion(database,workspace,id,user,'before-revert-'+input.operationId,'原位恢复前：'+version.name.slice(0,110),encodeOfficeBackup(snapshot.archive,snapshot.state));
  const media=remapOfficeMedia(state,resource.kind,parsed.media??[]);
  for(const entry of media.media){database.attachments.put('cloud:'+workspace,workspace,entry.id,entry.name,entry.mime,Buffer.from(entry.data,'base64'));database.db.prepare('INSERT INTO cloud_asset_owners VALUES (?,?,?,?)').run(workspace,entry.id,id,user);}
  const generation='epoch-'+randomUUID(),now=Date.now();database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(media.state),'cloud:'+workspace,id);database.db.prepare('UPDATE cloud_office_resources SET generation=?,updated=? WHERE workspace=? AND id=?').run(generation,now,workspace,id);
  database.db.prepare('INSERT INTO cloud_office_restores(workspace,user,operation,sha256,resource,created) VALUES (?,?,?,?,?,?)').run(workspace,user,input.operationId,signature,id,now);
  return {id,generation,repeated:false};
 }).immediate();
}
