import {presentationObjectIdentity,presentationFieldRevision,presentationObjectRevision,presentationArrangeRevision,presentationPaintOrder,arrangePresentationObjects,validPresentationField,PRESENTATION_FIELDS,PRESENTATION_OBJECT_TYPES,PRESENTATION_TEXT_FIELDS,type PresentationFieldOperation,type PresentationObjectOperation} from '../../../../common/nbstore/src/presentation-field';
import {validOfficeAssetId} from '../../../../common/nbstore/src/office-image';
import {readResourceMedia} from './office-media';
import {MAX_SLIDE_OBJECTS} from './office-presentation';
import * as Y from 'yjs';
import {createHash} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {changePresentationSlide,presentationSlides,createPresentationObject} from '../../../../common/nbstore/src/office-presentation';
import {structureRevision} from '../../../../common/nbstore/src/spreadsheet-operation';
import {validatePresentationUpdate} from './office-presentation';
export class PresentationOperationError extends Error {constructor(message:string,readonly status=400){super(message);}}
export function submitPresentationOperation(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:PresentationFieldOperation|PresentationObjectOperation|{operationId:string;action:'add'|'duplicate'|'delete'|'up'|'down';slide:string;base:string}):{update:Uint8Array;selected:string;object?:string;repeated:boolean}{
 if(input?.action==='setField')return submitPresentationField(database,workspace,id,user,input);
 if(input?.action==='addObject'||input?.action==='deleteObject'||input?.action==='arrangeObject')return submitPresentationObject(database,workspace,id,user,input);
 if(!input||!['add','duplicate','delete','up','down'].includes(input.action)||typeof input.slide!=='string'||input.slide.length>80||typeof input.base!=='string'||input.base.length>131072||typeof input.operationId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(input.operationId))throw new PresentationOperationError('幻灯片操作参数无效');
 const signature=createHash('sha256').update(JSON.stringify([user,input.action,input.slide,input.base])).digest('hex');
 return database.db.transaction(()=>{
  const scope='cloud:'+workspace,current=database.load(scope,id);if(!current)throw new PresentationOperationError('演示文稿不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('presentationOperations'),previous=receipts.get(input.operationId);
   if(previous){const [saved,selected]=previous.split(':');if(saved!==signature)throw new PresentationOperationError('操作编号已用于其他提交',409);return {update:current,selected,repeated:true};}
   if(receipts.size>=10000)throw new PresentationOperationError('演示操作记录已达上限，请备份后恢复为新文件',409);
   if(structureRevision(doc)!==input.base)throw new PresentationOperationError('演示文稿已更新，请重试并核对幻灯片顺序',409);
   if(!presentationSlides(doc).some(slide=>slide.id===input.slide))throw new PresentationOperationError('幻灯片已不存在',409);
   let selected:string;try{selected=changePresentationSlide(doc,input.action,input.slide);}catch(error){throw new PresentationOperationError(error instanceof Error?error.message:'幻灯片操作失败');}
   receipts.set(input.operationId,signature+':'+selected);const update=Y.encodeStateAsUpdate(doc);try{validatePresentationUpdate(null,update);}catch(error){throw new PresentationOperationError(error instanceof Error?error.message:'演示模型无效');}
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(update),scope,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update,selected,repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}

function submitPresentationField(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:PresentationFieldOperation){
 if(typeof input.operationId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(input.operationId)||typeof input.slide!=='string'||input.slide.length>80||typeof input.object!=='string'||!/^\d+-\d+$/.test(input.object)||typeof input.base!=='string'||input.base.length>700000||!validPresentationField(input.field,input.value))throw new PresentationOperationError('演示对象操作参数无效');
 const signature=createHash('sha256').update(JSON.stringify([user,input.slide,input.object,input.field,input.value,input.base])).digest('hex');
 return database.db.transaction(()=>{
  const scope='cloud:'+workspace,current=database.load(scope,id);if(!current)throw new PresentationOperationError('演示文稿不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('presentationOperations'),previous=receipts.get(input.operationId);
   if(previous){if(previous!==signature+':'+input.slide)throw new PresentationOperationError('操作编号已用于其他提交',409);return {update:current,selected:input.slide,repeated:true};}
   if(receipts.size>=10000)throw new PresentationOperationError('演示操作记录已达上限，请备份后恢复为新文件',409);
   const slide=presentationSlides(doc).find(slide=>slide.id===input.slide),object=slide?.objects.toArray().find(object=>presentationObjectIdentity(object)===input.object);
   if(!object||presentationFieldRevision(object,input.field)!==input.base)throw new PresentationOperationError('对象字段已被修改或删除，本人输入已保留，请核对最新内容再提交',409);
   if((input.field==='text'||PRESENTATION_TEXT_FIELDS.includes(input.field as never))&&object.get('type')!=='text')throw new PresentationOperationError('此对象不是文本框');
   object.set(input.field,input.value);doc.getMap('presentationFields').set(input.object+':'+input.field,input.operationId);receipts.set(input.operationId,signature+':'+input.slide);
   const update=Y.encodeStateAsUpdate(doc);try{validatePresentationUpdate(null,update);}catch(error){throw new PresentationOperationError(error instanceof Error?error.message:'演示模型无效');}
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(update),scope,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update,selected:input.slide,repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}

function submitPresentationObject(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:PresentationObjectOperation){
 const add=input.action==='addObject',arrange=input.action==='arrangeObject',geometry=add?input.geometry??{}:{};
 if(typeof input.operationId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(input.operationId)||typeof input.slide!=='string'||!/^[a-zA-Z0-9-]{1,80}$/.test(input.slide)||Object.keys(input).some(key=>!(add?['action','operationId','slide','type','assetId','geometry']:arrange?['action','operationId','slide','object','to','base']:['action','operationId','slide','object','base']).includes(key)))throw new PresentationOperationError('演示对象操作参数无效');
 if(arrange&&!['front','back','forward','backward'].includes(input.to))throw new PresentationOperationError('演示对象层次操作无效');
 if(add&&(!PRESENTATION_OBJECT_TYPES.includes(input.type)||(input.type==='image')!==(input.assetId!==undefined)||input.assetId!==undefined&&!validOfficeAssetId(input.assetId)||!geometry||typeof geometry!=='object'||Array.isArray(geometry)||Object.entries(geometry).some(([key,value])=>!['x','y','width','height'].includes(key)||!validPresentationField(key,value))))throw new PresentationOperationError('演示对象类型、图片或位置无效');
 if(!add&&(typeof input.object!=='string'||!/^\d+-\d+$/.test(input.object)||typeof input.base!=='string'||input.base.length>700000))throw new PresentationOperationError('演示对象操作参数无效');
 const signature=createHash('sha256').update(JSON.stringify(add?[user,input.action,input.slide,input.type,input.assetId??null,geometry]:arrange?[user,input.action,input.slide,input.object,input.to,input.base]:[user,input.action,input.slide,input.object,input.base])).digest('hex');
 return database.db.transaction(()=>{
  const scope='cloud:'+workspace,current=database.load(scope,id);if(!current)throw new PresentationOperationError('演示文稿不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('presentationOperations'),previous=receipts.get(input.operationId);
   // Receipts store the created object so a lost response can be retried without a second copy.
   if(previous){const [saved,object]=previous.split(':');if(saved!==signature)throw new PresentationOperationError('操作编号已用于其他提交',409);return {update:current,selected:input.slide,object,repeated:true};}
   if(receipts.size>=10000)throw new PresentationOperationError('演示操作记录已达上限，请备份后恢复为新文件',409);
   const slide=presentationSlides(doc).find(slide=>slide.id===input.slide);if(!slide)throw new PresentationOperationError('幻灯片已不存在，请重新选择',409);
   let object:string;
   if(add){
    if(slide.objects.length>=MAX_SLIDE_OBJECTS)throw new PresentationOperationError('幻灯片对象超过 1 万',413);
    const value=createPresentationObject(input.type);if(input.assetId)value.set('assetId',input.assetId);for(const [key,number]of Object.entries(geometry))value.set(key,number);
    // Once a slide has explicit layers, ranks are re-densified and the new object goes on top so `z` stays bounded.
    const existing=slide.objects.toArray();if(existing.some(item=>item.has('z'))){presentationPaintOrder(existing).forEach((entry,rank)=>{if(entry.object.get('z')!==rank)entry.object.set('z',rank);});value.set('z',existing.length);}
    slide.objects.push([value]);object=presentationObjectIdentity(value);
   }else if(arrange){
    const objects=slide.objects.toArray(),index=objects.findIndex(value=>presentationObjectIdentity(value)===input.object);if(index<0)throw new PresentationOperationError('对象已被其他成员删除',409);
    if(presentationArrangeRevision(objects)!==input.base)throw new PresentationOperationError('幻灯片对象层次已被其他成员调整，请核对后重试',409);
    for(const [at,rank]of arrangePresentationObjects(objects,index,input.to))if(objects[at].get('z')!==rank)objects[at].set('z',rank);object=input.object;
   }else{
    const index=slide.objects.toArray().findIndex(value=>presentationObjectIdentity(value)===input.object);if(index<0)throw new PresentationOperationError('对象已被其他成员删除',409);
    if(presentationObjectRevision(slide.objects.get(index))!==input.base)throw new PresentationOperationError('对象已被其他成员修改，未删除。请核对最新内容后再决定',409);
    slide.objects.delete(index,1);object=input.object;if(doc.share.has('presentationFields')){const fields=doc.getMap('presentationFields');for(const field of PRESENTATION_FIELDS)fields.delete(object+':'+field);}
   }
   receipts.set(input.operationId,signature+':'+object);const update=Y.encodeStateAsUpdate(doc);
   try{validatePresentationUpdate(null,update);if(add&&input.assetId)readResourceMedia(database,workspace,id,'presentation',update);}catch(error){throw new PresentationOperationError(error instanceof Error?error.message:'演示模型无效');}
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(update),scope,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update,selected:input.slide,object,repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}
