import * as Y from 'yjs';
import {wordBlockTexts,type WordBlock} from './office-word-table';
import {workbookSheets} from './office-workbook';
import {sheetColumnIds,sheetRowIdentity} from './sheet-columns';
import {presentationSlides} from './office-presentation';
import {presentationObjectIdentity} from './presentation-field';
import {pdfAnnotationId} from './office-pdf';
export type OfficeCommentAnchor={kind:'text';start:number[];end:number[];quote:string}|{kind:'cell';sheet:string;row:string;column:string;label:string}|{kind:'presentation';slide:string;object:string}|{kind:'pdf';annotation:string;page:number};
const record=(value:unknown):value is Record<string,unknown>=>!!value&&typeof value==='object'&&!Array.isArray(value);
const exact=(value:Record<string,unknown>,keys:string[])=>Object.keys(value).length===keys.length&&keys.every(key=>Object.hasOwn(value,key));
export function validOfficeCommentAnchor(value:unknown):value is OfficeCommentAnchor{
 if(!record(value))return false;const string=(key:string,max=160)=>typeof value[key]==='string'&&String(value[key]).length>0&&String(value[key]).length<=max;
 if(value.kind==='text')return exact(value,['kind','start','end','quote'])&&['start','end'].every(key=>Array.isArray(value[key])&&value[key].length>0&&value[key].length<=256&&value[key].every((n:unknown)=>Number.isInteger(n)&&Number(n)>=0&&Number(n)<=255))&&string('quote',500);
 if(value.kind==='cell')return exact(value,['kind','sheet','row','column','label'])&&['sheet','row','column','label'].every(key=>string(key));
 if(value.kind==='presentation')return exact(value,['kind','slide','object'])&&string('slide')&&string('object');
 return value.kind==='pdf'&&exact(value,['kind','annotation','page'])&&string('annotation')&&Number.isSafeInteger(value.page)&&Number(value.page)>0&&Number(value.page)<=100000;
}
export function captureOfficeTextAnchor(first:Y.Text,start:number,last:Y.Text,end:number,quote:string):OfficeCommentAnchor{
 return {kind:'text',start:Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(first,start,0))),end:Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(last,end,-1))),quote:quote.slice(0,500)};
}
export function officeAnchorTarget(doc:Y.Doc,anchor:OfficeCommentAnchor){
 try{
  if(anchor.kind==='text'){
   const allowed=doc.share.has('paragraphs')?wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray()):doc.share.has('text')?[doc.getText('text')]:[];
   const start=Y.createAbsolutePositionFromRelativePosition(Y.decodeRelativePosition(Uint8Array.from(anchor.start)),doc),end=Y.createAbsolutePositionFromRelativePosition(Y.decodeRelativePosition(Uint8Array.from(anchor.end)),doc);if(!start||!end||!(start.type instanceof Y.Text)||!(end.type instanceof Y.Text))return;
   const a=allowed.indexOf(start.type),b=allowed.indexOf(end.type);if(a<0||b<a||a===b&&end.index<=start.index)return;
   return {kind:'text' as const,first:start.type,last:end.type,start:start.index,end:end.index,label:anchor.quote};
  }
  if(anchor.kind==='cell'){
   if(!doc.share.has('sheet'))return;const sheet=workbookSheets(doc).find(sheet=>sheet.id===anchor.sheet),row=sheet?.sheet.toArray().findIndex(row=>sheetRowIdentity(row)===anchor.row)??-1,column=sheet?sheetColumnIds(sheet.sheet).indexOf(anchor.column):-1;if(!sheet||row<0||column<0)return;return {kind:'cell' as const,sheet:sheet.id,row,column,label:sheet.name+' · '+anchor.label};
  }
  if(anchor.kind==='presentation'){
   if(!doc.share.has('objects'))return;const slide=presentationSlides(doc).find(slide=>slide.id===anchor.slide),index=slide?.objects.toArray().findIndex(object=>presentationObjectIdentity(object)===anchor.object)??-1;if(!slide||index<0)return;return {kind:'presentation' as const,slide:slide.id,index,label:'幻灯片对象 '+(index+1)};
  }
  if(!doc.share.has('annotations'))return;const annotation=doc.getArray<Y.Map<unknown>>('annotations').toArray().find(value=>pdfAnnotationId(value)===anchor.annotation);if(!annotation)return;return {kind:'pdf' as const,annotation,page:Number(annotation.get('page')),label:'PDF 第 '+annotation.get('page')+' 页批注'};
 }catch{return;}
}
