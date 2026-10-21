import * as Y from 'yjs';
import {createHash,randomUUID} from 'node:crypto';
import type {WorkspaceBackup} from '../../../core/src/backup';
import {BlockStore} from '../../../../editor/framework/src/index';
import {remapAssets,remapReferences} from '../../../../editor/blocks/src/document-format';
import {prepareWorkspaceArchive} from './workspace-archive-model';
import {archiveBinary} from './workspace-archive';
import type {CloudWorkspaceArchive} from '../../../../common/nbstore/src/workspace-archive';
/** Runs only in the isolated archive parser. The public v2 format stays unchanged. */
export function preparePersonalWorkspaceArchive(bytes:Uint8Array,target:string){
 if(bytes.length>50*1024*1024)throw new Error('完整备份超过 50 MB');const backup=JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(bytes)) as WorkspaceBackup;
 if(!backup||backup.type!=='mydoc-workspace'||backup.version!==2||!backup.office||typeof backup.name!=='string'||!backup.name.trim()||typeof backup.workspaceId!=='string'||!Array.isArray(backup.documents)||!Array.isArray(backup.folders))throw new Error('原子恢复需要个人工作区 v2 备份');
 for(const list of [backup.documents,backup.folders,backup.trash??[],backup.tags??[],backup.collections??[],backup.assets??[],backup.templates??[],backup.officeOrganization??[]])if(!Array.isArray(list)||list.length>10000)throw new Error('个人备份列表无效或超限');
 if(backup.documents.length+(backup.trash?.length??0)>10000)throw new Error('个人文档数量超限');const decode=(value:string,max=16*1024*1024)=>{if(typeof value!=='string'||value.length>Math.ceil(max/3)*4)throw new Error('备份二进制长度无效');const result=Buffer.from(value,'base64');if(result.toString('base64')!==value||result.length>max)throw new Error('备份二进制数据无效');return result;};
 for(const entry of [...backup.documents,...(backup.trash??[])])if(!entry||!entry.meta||typeof entry.meta.id!=='string'||typeof entry.meta.title!=='string'||!Number.isFinite(entry.meta.createdAt))throw new Error('个人备份文档信息无效');
 const noteIds=[...backup.documents,...(backup.trash??[])].map(entry=>entry.meta.id);if(new Set(noteIds).size!==noteIds.length)throw new Error('个人文档编号重复');const assetKeys=new Set<string>();for(const asset of backup.assets??[]){if(!asset||typeof asset.id!=='string'||assetKeys.has(asset.id)||decode(asset.data,20*1024*1024).length!==asset.size)throw new Error('个人备份附件大小或编号无效');assetKeys.add(asset.id);}
 
 for(const list of [backup.folders,backup.tags??[],backup.collections??[],backup.templates??[],backup.officeOrganization??[]]){const seen=new Set<string>();for(const entry of list){if(!entry||typeof entry.id!=='string'||!entry.id||seen.has(entry.id))throw new Error('个人备份目录标识无效或重复');seen.add(entry.id);}}
 const office=JSON.parse(Buffer.from(prepareWorkspaceArchive(Buffer.from(JSON.stringify(backup.office)))).toString()) as CloudWorkspaceArchive,officeIds=new Set(office.office.map(file=>file.id));
 if(noteIds.some(id=>officeIds.has(id)))throw new Error('普通文档与办公标识冲突');
 if(office.documents.some(entry=>entry.id!=='workspace:'+office.workspace.id&&!officeIds.has(entry.id)))throw new Error('个人办公备份包含非办公文档');
 const source=backup.workspaceId??'local',entries=[...backup.documents,...(backup.trash??[])],ids=new Map<string,string>(entries.map(entry=>[entry.meta.id,randomUUID()]));for(const id of officeIds)ids.set(id,id);
 const assetIds=new Map((backup.assets??[]).map(asset=>[asset.id,randomUUID()]));const rewrite=(data:string)=>{const doc=new Y.Doc();let store:BlockStore|undefined;try{Y.applyUpdate(doc,decode(data));store=new BlockStore(doc);remapReferences(store,source,target,ids);remapAssets(store,assetIds);return Y.encodeStateAsUpdate(doc);}finally{store?.destroy();doc.destroy();}};
 const meta=new Y.Doc();try{
  for(const folder of backup.folders)meta.getMap('folders').set(folder.id,folder);for(const tag of backup.tags??[])meta.getMap('tags').set(tag.id,tag);
  for(const collection of backup.collections??[])meta.getMap('collections').set(collection.id,{...collection,documentIds:collection.documentIds.map(id=>ids.get(id)!)});
  for(const entry of backup.documents)meta.getMap('documents').set(ids.get(entry.meta.id)!,{...entry.meta,id:ids.get(entry.meta.id)!});
  for(const entry of backup.trash??[])meta.getMap('trash').set(ids.get(entry.meta.id)!,{meta:{...entry.meta,id:ids.get(entry.meta.id)!},deletedAt:entry.deletedAt,collectionIds:entry.collectionIds});
  for(const {id,...value}of backup.officeOrganization??[])meta.getMap('officeOrganization').set(id,value);
  for(const template of backup.templates??[])meta.getMap('documentTemplates').set(template.id,{...template,sourceWorkspaceId:template.sourceWorkspaceId===source?target:template.sourceWorkspaceId,sourceDocId:template.sourceWorkspaceId===source?ids.get(template.sourceDocId)??template.sourceDocId:template.sourceDocId,update:Buffer.from(rewrite(template.update)).toString('base64')});
  const archive:CloudWorkspaceArchive={...office,workspace:{id:target,name:backup.name},documents:[...office.documents.filter(entry=>officeIds.has(entry.id)),...entries.map(entry=>({id:'doc:'+ids.get(entry.meta.id)!,state:archiveBinary(rewrite(entry.update))})),{id:'workspace:'+target,state:archiveBinary(Y.encodeStateAsUpdate(meta))}],assets:[...office.assets,...(backup.assets??[]).map(asset=>({id:assetIds.get(asset.id)!,name:asset.name,mime:asset.mime,doc:null,state:archiveBinary(Buffer.from(asset.data,'base64'))}))]};
  return prepareWorkspaceArchive(Buffer.from(JSON.stringify(archive)),target);
 }finally{meta.destroy();}
}
