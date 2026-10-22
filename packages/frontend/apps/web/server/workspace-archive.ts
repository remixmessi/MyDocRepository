import {createHash,randomUUID} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {MAX_WORKSPACE_ARCHIVE_BYTES,type CloudWorkspaceArchive,type ArchiveBinary} from '../../../../common/nbstore/src/workspace-archive';
export const archiveBinary=(bytes:Uint8Array):ArchiveBinary=>({data:Buffer.from(bytes).toString('base64'),sha256:createHash('sha256').update(bytes).digest('hex')});
export class WorkspaceArchiveError extends Error{constructor(message:string,readonly status=400){super(message);}}
export function readWorkspaceArchive(database:SQLiteProjectDatabase,workspace:string,localName?:string):Uint8Array{
 return database.db.transaction(()=>{
  const info=localName?{name:localName}:database.db.prepare('SELECT name FROM cloud_workspaces WHERE id=?').get(workspace) as {name:string}|undefined;if(!info)throw new WorkspaceArchiveError('工作区不存在',404);
  const stats=database.db.prepare('SELECT count(*) count,coalesce(sum(length(data)),0) bytes,coalesce(max(length(data)),0) largest FROM documents WHERE scope=?').get('cloud:'+workspace) as {count:number;bytes:number;largest:number};
  const assetStats=database.db.prepare('SELECT count(*) count,coalesce(sum(coalesce(f.size,length(a.data))),0) bytes,coalesce(max(coalesce(f.size,length(a.data))),0) largest FROM assets a LEFT JOIN asset_files f ON f.scope=a.scope AND f.workspace=a.workspace AND f.id=a.id WHERE a.scope=? AND a.workspace=?').get('cloud:'+workspace,workspace) as {count:number;bytes:number;largest:number};
  if(stats.count>10001||assetStats.count>10000)throw new WorkspaceArchiveError('工作区超过完整备份数量上限');
  if(stats.largest>16*1024*1024||assetStats.largest>20*1024*1024||stats.bytes+assetStats.bytes>MAX_WORKSPACE_ARCHIVE_BYTES*3/4)throw new WorkspaceArchiveError('完整备份超过单文件或 50 MB 总量限制，请拆分工作区');
  const versionStats=database.db.prepare('SELECT count(*) count,coalesce(sum(length(data)),0) bytes FROM cloud_office_versions WHERE workspace=? AND deleted=0').get(workspace) as {count:number;bytes:number};if(versionStats.count>10000||stats.bytes+assetStats.bytes+versionStats.bytes>MAX_WORKSPACE_ARCHIVE_BYTES*3/4)throw new WorkspaceArchiveError('包含历史版本的完整备份超过 50 MB，请先下载并清理不需要的版本');
  const documents=database.db.prepare('SELECT id,data FROM documents WHERE scope=? ORDER BY id').all('cloud:'+workspace) as {id:string;data:Buffer}[],assets=database.db.prepare('SELECT a.id,a.name,a.mime,o.doc FROM assets a LEFT JOIN cloud_asset_owners o ON o.workspace=a.workspace AND o.id=a.id WHERE a.scope=? AND a.workspace=? ORDER BY a.id').all('cloud:'+workspace,workspace) as {id:string;name:string;mime:string;doc:string|null}[];
  if(documents.length>10001||assets.length>10000)throw new WorkspaceArchiveError('工作区超过完整备份数量上限');let size=0;
  for(const row of documents){size+=row.data.length;if(row.data.length>16*1024*1024)throw new WorkspaceArchiveError('单份协作状态超过 16 MB');}if(size>MAX_WORKSPACE_ARCHIVE_BYTES*3/4)throw new WorkspaceArchiveError('完整备份超过 50 MB，请拆分工作区');
  const binaryAssets=assets.map(row=>{const asset=database.attachments.get('cloud:'+workspace,workspace,row.id);if(!asset)throw new WorkspaceArchiveError('附件缺失，无法生成完整备份');size+=asset.data.length;if(size>MAX_WORKSPACE_ARCHIVE_BYTES*3/4)throw new WorkspaceArchiveError('完整备份超过 50 MB，请拆分工作区');return {...row,doc:row.doc??null,state:archiveBinary(asset.data)};});
  const office=database.db.prepare('SELECT id,kind,name,mime,source_format AS sourceFormat,original_asset_id AS originalAssetId,model_version AS modelVersion,created,updated,deleted FROM cloud_office_resources WHERE workspace=? ORDER BY id LIMIT 10001').all(workspace) as CloudWorkspaceArchive['office'];
  if(office.length>10000)throw new WorkspaceArchiveError('办公资源数量超限');
  const tombstones=(database.db.prepare('SELECT id FROM cloud_deleted WHERE workspace=? ORDER BY id LIMIT 10001').all(workspace) as {id:string}[]).map(row=>row.id);if(tombstones.length>10000)throw new WorkspaceArchiveError('永久删除标记数量超限');
  const versions=(database.db.prepare('SELECT id,resource,name,created,data,sha256 FROM cloud_office_versions WHERE workspace=? AND deleted=0 ORDER BY resource,created,id').all(workspace) as {id:string;resource:string;name:string;created:number;data:Buffer;sha256:string}[]).map(({data,sha256,...entry})=>{const state=archiveBinary(data);if(state.sha256!==sha256)throw new WorkspaceArchiveError('历史版本校验失败');return {...entry,state};});
  const archive:CloudWorkspaceArchive={type:'mydoc-team-workspace',version:2,versions,exportedAt:Date.now(),workspace:{id:workspace,name:info.name},documents:documents.map(row=>({id:row.id,state:archiveBinary(row.data)})),office,assets:binaryAssets,deleted:tombstones};
  const bytes=Buffer.from(JSON.stringify(archive));if(bytes.length>MAX_WORKSPACE_ARCHIVE_BYTES)throw new WorkspaceArchiveError('完整备份超过 50 MB，请拆分工作区');return bytes;
 }).immediate();
}
export function workspaceRestoreReceipt(database:SQLiteProjectDatabase,user:string,operation:string,sha256:string){
 const receipt=database.db.prepare('SELECT sha256,workspace FROM cloud_workspace_restores WHERE user=? AND operation=?').get(user,operation) as {sha256:string;workspace:string}|undefined;if(!receipt)return;
 if(receipt.sha256!==sha256)throw new WorkspaceArchiveError('恢复编号已用于另一份备份，请重新选择文件',409);
 const space=database.db.prepare('SELECT w.id,w.name,w.created AS createdAt,m.role FROM cloud_workspaces w JOIN cloud_members m ON m.workspace=w.id AND m.user=? WHERE w.id=?').get(user,receipt.workspace) as {id:string;name:string;createdAt:number;role:string}|undefined;if(!space||space.role!=='owner')throw new WorkspaceArchiveError('此前恢复的工作区已删除或权限已变化，请选择备份开始新的恢复',410);return space;
}
/** The caller validates and remaps every snapshot in an isolated process before this single write transaction. */
export function restoreWorkspaceArchive(database:SQLiteProjectDatabase,user:string,operation:string,sha256:string,archive:CloudWorkspaceArchive){
 return database.db.transaction(()=>{
  const existing=workspaceRestoreReceipt(database,user,operation,sha256);if(existing)return {...existing,repeated:true};
  const id=archive.workspace.id,name=archive.workspace.name.slice(0,156)+'（恢复）',createdAt=Date.now(),scope='cloud:'+id;
  if(database.db.prepare('SELECT 1 FROM cloud_workspaces WHERE id=?').get(id))throw new WorkspaceArchiveError('恢复工作区编号冲突，请重试',409);
  database.db.prepare('INSERT INTO cloud_workspaces VALUES (?,?,?)').run(id,name,createdAt);database.db.prepare('INSERT INTO cloud_members VALUES (?,?,?)').run(id,user,'owner');
  for(const row of archive.documents)database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run(scope,row.id,Buffer.from(row.state.data,'base64'));
  for(const asset of archive.assets){database.attachments.put(scope,id,asset.id,asset.name,asset.mime,Buffer.from(asset.state.data,'base64'));database.db.prepare('INSERT INTO cloud_asset_owners VALUES (?,?,?,?)').run(id,asset.id,asset.doc,user);}
  for(const r of archive.office)database.db.prepare('INSERT INTO cloud_office_resources(workspace,id,kind,name,mime,source_format,original_asset_id,model_version,generation,created_by,created,updated,deleted) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)').run(id,r.id,r.kind,r.name,r.mime,r.sourceFormat,r.originalAssetId,1,randomUUID(),user,r.created,r.updated,r.deleted);
  for(const v of archive.versions??[])database.db.prepare('INSERT INTO cloud_office_versions(workspace,resource,id,name,user,operation,created,data,sha256,deleted) VALUES (?,?,?,?,?,?,?,?,?,0)').run(id,v.resource,v.id,v.name,user,'archive-'+randomUUID(),v.created,Buffer.from(v.state.data,'base64'),v.state.sha256);
  for(const doc of archive.deleted)database.db.prepare('INSERT INTO cloud_deleted VALUES (?,?)').run(id,doc);
  database.db.prepare('INSERT INTO cloud_workspace_restores VALUES (?,?,?,?,?)').run(user,operation,sha256,id,createdAt);
  return {id,name,createdAt,role:'owner' as const,repeated:false};
 }).immediate();
}

/** Restores only an office namespace. The local registry and notes use the ordinary SQLite storage boundary. */
export function restoreLocalOfficeArchive(database:SQLiteProjectDatabase,workspace:string,user:string,archive:CloudWorkspaceArchive){
 return database.db.transaction(()=>{
  if(database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=?').get(workspace))throw new WorkspaceArchiveError('恢复目标已有办公文件，不能覆盖',409);
  if(archive.workspace.id!==workspace||archive.documents.some(row=>row.id!=='workspace:'+workspace&&!archive.office.some(r=>r.id===row.id)))throw new WorkspaceArchiveError('个人办公备份含有非办公数据');
  const scope='cloud:'+workspace;
  for(const row of archive.documents)if(row.id!=='workspace:'+workspace)database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run(scope,row.id,Buffer.from(row.state.data,'base64'));
  for(const asset of archive.assets){database.attachments.put(scope,workspace,asset.id,asset.name,asset.mime,Buffer.from(asset.state.data,'base64'));database.db.prepare('INSERT INTO cloud_asset_owners VALUES (?,?,?,?)').run(workspace,asset.id,asset.doc,user);}
  for(const r of archive.office)database.db.prepare('INSERT INTO cloud_office_resources(workspace,id,kind,name,mime,source_format,original_asset_id,model_version,generation,created_by,created,updated,deleted) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)').run(workspace,r.id,r.kind,r.name,r.mime,r.sourceFormat,r.originalAssetId,1,randomUUID(),user,r.created,r.updated,r.deleted);
  for(const v of archive.versions??[])database.db.prepare('INSERT INTO cloud_office_versions(workspace,resource,id,name,user,operation,created,data,sha256,deleted) VALUES (?,?,?,?,?,?,?,?,?,0)').run(workspace,v.resource,v.id,v.name,user,'archive-'+randomUUID(),v.created,Buffer.from(v.state.data,'base64'),v.state.sha256);
  for(const id of archive.deleted)database.db.prepare('INSERT INTO cloud_deleted VALUES (?,?)').run(workspace,id);
 }).immediate();
}
