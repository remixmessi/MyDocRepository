import {validateOfficeMedia,readResourceMedia,remapOfficeMedia,type OfficeMedia} from './office-media';
import {createHash,randomUUID} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {officeKindFromFilename,type OfficeKind} from '../../../../common/nbstore/src/office';

export const MAX_OFFICE_BACKUP_BYTES=50*1024*1024;
export type OfficeBackup={type:'mydoc-office-resource';version:1|2;media?:OfficeMedia[];exportedAt:number;resource:{name:string;kind:OfficeKind;mime:string;sourceFormat:string;modelVersion:1};state:{data:string;sha256:string};original:null|{name:string;mime:string;data:string;sha256:string}};
const digest=(bytes:Uint8Array)=>createHash('sha256').update(bytes).digest('hex');
function object(value:unknown):value is Record<string,unknown>{return !!value&&typeof value==='object'&&!Array.isArray(value);}
function fields(value:Record<string,unknown>,keys:string[]){return Object.keys(value).length===keys.length&&keys.every(key=>Object.hasOwn(value,key));}
function label(value:unknown,max:number):value is string{return typeof value==='string'&&!!value.trim()&&value.length<=max&&!/[\x00-\x1f]/.test(value);}
function decode(value:unknown,max:number){
 if(!object(value)||!label(value.sha256,64)||!/^[a-f0-9]{64}$/.test(value.sha256)||typeof value.data!=='string'||value.data.length>Math.ceil(max/3)*4||! /^[A-Za-z0-9+/]*={0,2}$/.test(value.data))throw new Error('备份二进制信息无效');
 const bytes=Buffer.from(value.data,'base64');if(bytes.length>max||bytes.toString('base64')!==value.data||digest(bytes)!==value.sha256)throw new Error('备份校验失败，文件可能已损坏');return bytes;
}
export function parseOfficeBackup(bytes:Uint8Array){
 if(bytes.length>MAX_OFFICE_BACKUP_BYTES)throw new Error('办公文件备份超过 50 MB');
 const value:unknown=JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(bytes));
 if(!object(value)||!fields(value,['type','version','exportedAt','resource','state','original',...(value.version===2?['media']:[])])||value.type!=='mydoc-office-resource'||(value.version!==1&&value.version!==2)||!Number.isSafeInteger(value.exportedAt)||Number(value.exportedAt)<0||!object(value.resource)||!fields(value.resource,['name','kind','mime','sourceFormat','modelVersion']))throw new Error('不是支持的办公文件备份');
 const r=value.resource;if(!label(r.name,240)||!label(r.mime,200)||!label(r.sourceFormat,20)||r.modelVersion!==1)throw new Error('备份资源信息无效');
 const format=officeKindFromFilename('file.'+r.sourceFormat);if(!format||format.kind!==r.kind||format.sourceFormat!==r.sourceFormat)throw new Error('备份资源格式不匹配');
 if(!object(value.state)||!fields(value.state,['data','sha256']))throw new Error('备份状态无效');const state=decode(value.state,16*1024*1024);
 let original:Buffer|null=null;if(value.original!==null){if(!object(value.original)||!fields(value.original,['name','mime','data','sha256'])||!label(value.original.name,240)||!label(value.original.mime,200))throw new Error('备份原件信息无效');original=decode(value.original,20*1024*1024);}
 if(r.kind==='pdf'&&(!original||original.subarray(0,5).toString()!=='%PDF-'))throw new Error('PDF 备份缺少原件');
 validateOfficeMedia((value.version===2?value.media:[]) as OfficeMedia[]);
 return {archive:value as unknown as OfficeBackup,state,original};
}
export function readOfficeBackup(database:SQLiteProjectDatabase,workspace:string,id:string){
 return database.db.transaction(()=>{
  const resource=database.db.prepare('SELECT name,kind,mime,source_format AS sourceFormat,model_version AS modelVersion,original_asset_id AS asset FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id) as OfficeBackup['resource']&{asset:string|null}|undefined;
  if(!resource)throw new Error('办公文件不存在或已删除');const state=database.load('cloud:'+workspace,id);if(!state||state.length>16*1024*1024)throw new Error('协作状态缺失或超过 16 MB');
  let original:OfficeBackup['original']=null;if(resource.asset){const binding=database.db.prepare('SELECT doc FROM cloud_asset_owners WHERE workspace=? AND id=?').get(workspace,resource.asset) as {doc:string}|undefined;if(binding?.doc!==id)throw new Error('原件归属无效');const asset=database.attachments.get('cloud:'+workspace,workspace,resource.asset);if(!asset)throw new Error('原件缺失，无法生成备份');original={name:asset.name,mime:asset.mime,data:Buffer.from(asset.data).toString('base64'),sha256:digest(asset.data)};}
  const media=readResourceMedia(database,workspace,id,resource.kind,state);const {asset:_,...meta}=resource;
  return {archive:{type:'mydoc-office-resource',version:2,media,exportedAt:Date.now(),resource:meta,state:{data:Buffer.from(state).toString('base64'),sha256:digest(state)},original} as OfficeBackup,state:Buffer.from(state)};
 }).immediate();
}
export function encodeOfficeBackup(archive:OfficeBackup,state:Uint8Array){
 const result={...archive,state:{data:Buffer.from(state).toString('base64'),sha256:digest(state)}};const bytes=Buffer.from(JSON.stringify(result));if(bytes.length>MAX_OFFICE_BACKUP_BYTES)throw new Error('办公文件备份超过 50 MB');parseOfficeBackup(bytes);return bytes;
}
export class OfficeRestoreError extends Error{constructor(message:string,readonly status:number){super(message);}}
export function restoreOfficeBackup(database:SQLiteProjectDatabase,workspace:string,user:string,parsed:ReturnType<typeof parseOfficeBackup>,state:Uint8Array,operation?:{id:string;sha256:string}){
 const {archive,original}=parsed,remapped=remapOfficeMedia(state,parsed.archive.resource.kind,parsed.archive.media??[]),id='doc:'+randomUUID(),assetId=original?randomUUID():null,generation=randomUUID(),now=Date.now(),r=archive.resource;
 return database.db.transaction(()=>{
  if(operation){
   const receipt=database.db.prepare('SELECT sha256,resource FROM cloud_office_restores WHERE workspace=? AND user=? AND operation=?').get(workspace,user,operation.id) as {sha256:string;resource:string}|undefined;
   if(receipt){
    if(receipt.sha256!==operation.sha256)throw new OfficeRestoreError('恢复编号已用于另一份备份，请重新选择文件',409);
    const resource=database.db.prepare('SELECT id,kind,name,mime,source_format AS sourceFormat,original_asset_id AS originalAssetId,model_version AS modelVersion,generation,created_by AS createdBy,created,updated,deleted FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,receipt.resource);
    if(!resource)throw new OfficeRestoreError('此前恢复的文件已删除，请重新选择备份开始新的恢复',410);
    return resource as {id:string};
   }
  }
  if(original&&assetId){database.attachments.put('cloud:'+workspace,workspace,assetId,archive.original!.name,archive.original!.mime,original);database.db.prepare('INSERT INTO cloud_asset_owners(workspace,id,doc,user) VALUES (?,?,?,?)').run(workspace,assetId,id,user);}
  for(const media of remapped.media){database.attachments.put('cloud:'+workspace,workspace,media.id,media.name,media.mime,Buffer.from(media.data,'base64'));database.db.prepare('INSERT INTO cloud_asset_owners(workspace,id,doc,user) VALUES (?,?,?,?)').run(workspace,media.id,id,user);}
  database.db.prepare('INSERT INTO cloud_office_resources(workspace,id,kind,name,mime,source_format,original_asset_id,model_version,generation,created_by,created,updated,deleted) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,0)').run(workspace,id,r.kind,r.name,r.mime,r.sourceFormat,assetId,1,generation,user,now,now);
  database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:'+workspace,id,Buffer.from(remapped.state));
  if(operation)database.db.prepare('INSERT INTO cloud_office_restores(workspace,user,operation,sha256,resource,created) VALUES (?,?,?,?,?,?)').run(workspace,user,operation.id,operation.sha256,id,now);
  return {id,...r,originalAssetId:assetId,generation,createdBy:user,created:now,updated:now,deleted:0};
 }).immediate();
}
