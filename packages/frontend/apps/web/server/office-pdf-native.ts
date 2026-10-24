import {PDFDocument,PDFName,PDFDict,PDFArray,PDFRef,PDFNumber,PDFString,PDFHexString,PDFStream,PDFPage,pushGraphicsState,popGraphicsState,concatTransformationMatrix,drawObject,rgb} from 'pdf-lib';
import type {PdfAnnotation} from '../../../../common/nbstore/src/office-pdf';
/** A native annotation already present in the immutable original. `key` is stable because the original never changes. */
export type PdfNativeAnnotation={key:string;page:number;subtype:string;annotation?:PdfAnnotation;author?:string;reason?:string};
const IMPORTABLE:Record<string,PdfAnnotation['type']>={Text:'comment',Highlight:'highlight',Ink:'ink'};
const numbers=(value:unknown)=>value instanceof PDFArray?value.asArray().map(item=>item instanceof PDFNumber?item.asNumber():NaN):[];
const text=(value:unknown)=>value instanceof PDFString||value instanceof PDFHexString?value.decodeText():'';
function annotations(page:PDFPage){const list=page.node.Annots(),result:{key:string;dict:PDFDict;ref?:PDFRef;index:number}[]=[];if(!list)return result;for(let index=0;index<list.size();index++){const raw=list.get(index),dict=list.lookup(index);if(dict instanceof PDFDict)result.push({key:raw instanceof PDFRef?raw.toString().replace(/\s+/g,'-'):'inline',dict,ref:raw instanceof PDFRef?raw:undefined,index});}return result;}
function nativeKey(page:number,entry:{key:string;index:number}){return entry.key==='inline'?`p${page}-i${entry.index}`:'r'+entry.key;}
/** Convert supported native annotations into model annotations, explaining why others stay read-only. */
export function readNativeAnnotations(pdf:PDFDocument){
 const result:PdfNativeAnnotation[]=[];
 pdf.getPages().forEach((page,i)=>{const box=page.getCropBox();for(const entry of annotations(page)){
  const subtype=(entry.dict.get(PDFName.of('Subtype')) as PDFName|undefined)?.decodeText()??'';if(['Popup','Widget','Link'].includes(subtype))continue;
  const key=nativeKey(i+1,entry),rect=numbers(entry.dict.lookup(PDFName.of('Rect'))),content=text(entry.dict.lookup(PDFName.of('Contents'))).slice(0,20000),author=text(entry.dict.lookup(PDFName.of('T'))).slice(0,100)||undefined;
  const base={key,page:i+1,subtype,...(author?{author}:{})},type=IMPORTABLE[subtype];
  if(!type){result.push({...base,reason:'此类原生批注仅可保留或按外观扁平化'});continue;}
  if(rect.length!==4||rect.some(n=>!Number.isFinite(n))){result.push({...base,reason:'批注缺少有效位置'});continue;}
  const [x1,y1,x2,y2]=rect,x=Math.min(x1,x2),y=Math.min(y1,y2),width=Math.abs(x2-x1),height=Math.abs(y2-y1);
  let annotation:PdfAnnotation;
  if(type==='ink'){const strokes=entry.dict.lookup(PDFName.of('InkList'));const paths=strokes instanceof PDFArray?strokes.asArray().map(stroke=>numbers(stroke instanceof PDFRef?pdf.context.lookup(stroke):stroke)):[];// Several strokes (a signature) keep their separate paths through `strokes`.
   const flat=paths.flat();if(!paths.length||paths.length>100||paths.some(path=>path.length<4||path.length%2)||flat.length>2000||flat.some(n=>!Number.isFinite(n))){result.push({...base,reason:'超长手绘暂不能转为可编辑批注'});continue;}annotation={page:i+1,type,x,y,width:Math.max(width,1),height:Math.max(height,1),content,points:flat,...(paths.length>1?{strokes:paths.map(path=>path.length/2)}:{})};}
  else if(type==='highlight')annotation={page:i+1,type,x,y,width:Math.max(width,1),height:Math.max(height,1),content};
  else annotation={page:i+1,type,x:x1,y:y1,content};
  const w=annotation.width??20,h=annotation.height??20;
  if(annotation.x<box.x||annotation.y<box.y||annotation.x+w>box.x+box.width||annotation.y+h>box.y+box.height){result.push({...base,reason:'批注超出页面裁剪区域'});continue;}
  if(annotation.points?.some((n,index)=>index%2===0?n<x||n>x+w:n<y||n>y+h)){result.push({...base,reason:'手绘路径超出批注区域'});continue;}
  result.push({...base,annotation});
 }});
 if(result.length>5000)throw new Error('原件批注超过 5000 个，暂不能导入');return result;
}
/** Drop originals that were imported into the collaborative model (the model is now their source of truth). */
export function removeImportedNatives(pdf:PDFDocument,imported:Set<string>){
 if(!imported.size)return;pdf.getPages().forEach((page,i)=>{const list=page.node.Annots();if(!list)return;const entries=annotations(page),drop=new Set<number>(),popups=new Set<string>();
  for(const entry of entries)if(imported.has(nativeKey(i+1,entry))){drop.add(entry.index);const popup=entry.dict.get(PDFName.of('Popup'));if(popup instanceof PDFRef)popups.add(popup.toString());}
  for(const entry of entries)if(entry.ref&&popups.has(entry.ref.toString()))drop.add(entry.index);
  for(const index of [...drop].sort((a,b)=>b-a))list.remove(index);
 });
}
/** Draw each remaining non-widget annotation from its normal appearance stream, then let the caller remove Annots. */
export function flattenNativeAppearances(pdf:PDFDocument){
 pdf.getPages().forEach((page,i)=>{for(const entry of annotations(page)){
  const subtype=(entry.dict.get(PDFName.of('Subtype')) as PDFName|undefined)?.decodeText()??'',flags=(entry.dict.lookup(PDFName.of('F')) as PDFNumber|undefined)?.asNumber()??0;
  if(['Popup','Widget','Link'].includes(subtype)||flags&2||flags&32)continue;
  let appearance=entry.dict.lookup(PDFName.of('AP'));appearance=appearance instanceof PDFDict?appearance.get(PDFName.of('N')):undefined;
  if(appearance instanceof PDFRef&&pdf.context.lookup(appearance) instanceof PDFDict){const states=pdf.context.lookup(appearance) as PDFDict,state=entry.dict.get(PDFName.of('AS'));appearance=state instanceof PDFName?states.get(state):undefined;}
  else if(appearance instanceof PDFDict&&!(appearance instanceof PDFStream)){const state=entry.dict.get(PDFName.of('AS'));appearance=state instanceof PDFName?appearance.get(state):undefined;}
  const rect=numbers(entry.dict.lookup(PDFName.of('Rect')));
  if(!(appearance instanceof PDFRef)||!(pdf.context.lookup(appearance) instanceof PDFStream)||rect.length!==4){
   // Without an appearance stream, draw the few types whose geometry fully defines how they look.
   const color=numbers(entry.dict.lookup(PDFName.of('C'))),stroke=color.length===3&&color.every(n=>n>=0&&n<=1)?rgb(color[0],color[1],color[2]):rgb(.15,.35,.9);
   if(subtype==='Highlight'&&rect.length===4){const [x1,y1,x2,y2]=rect;page.drawRectangle({x:Math.min(x1,x2),y:Math.min(y1,y2),width:Math.abs(x2-x1),height:Math.abs(y2-y1),color:color.length===3?stroke:rgb(1,.85,0),opacity:.35});continue;}
   if(subtype==='Ink'){const strokes=entry.dict.lookup(PDFName.of('InkList')),paths=strokes instanceof PDFArray?strokes.asArray().map(path=>numbers(path instanceof PDFRef?pdf.context.lookup(path):path)):[];if(paths.length&&paths.every(path=>path.length>=4&&path.length%2===0&&path.every(Number.isFinite))){for(const path of paths)for(let n=2;n<path.length;n+=2)page.drawLine({start:{x:path[n-2],y:path[n-1]},end:{x:path[n],y:path[n+1]},color:stroke,thickness:2});continue;}}
   if((subtype==='Square'||subtype==='Circle')&&rect.length===4){const [x1,y1,x2,y2]=rect,x=Math.min(x1,x2),y=Math.min(y1,y2),width=Math.abs(x2-x1),height=Math.abs(y2-y1);if(subtype==='Square')page.drawRectangle({x,y,width,height,borderColor:stroke,borderWidth:1});else page.drawEllipse({x:x+width/2,y:y+height/2,xScale:width/2,yScale:height/2,borderColor:stroke,borderWidth:1});continue;}
   throw new Error(`原件第 ${i+1} 页的 ${subtype||'未知'} 批注缺少外观，无法完整扁平化；请使用普通批注副本`);
  }
  const stream=pdf.context.lookup(appearance) as PDFStream,bbox=numbers(stream.dict.lookup(PDFName.of('BBox'))),matrix=stream.dict.has(PDFName.of('Matrix'))?numbers(stream.dict.lookup(PDFName.of('Matrix'))):[1,0,0,1,0,0];
  if(bbox.length!==4||matrix.length!==6||[...bbox,...matrix,...rect].some(n=>!Number.isFinite(n)))throw new Error(`原件第 ${i+1} 页批注外观无效，无法扁平化`);
  // ECMA/ISO 32000 12.5.5: map the transformed BBox onto Rect.
  const corners=[[bbox[0],bbox[1]],[bbox[2],bbox[1]],[bbox[0],bbox[3]],[bbox[2],bbox[3]]].map(([x,y])=>[matrix[0]*x+matrix[2]*y+matrix[4],matrix[1]*x+matrix[3]*y+matrix[5]]);
  const minX=Math.min(...corners.map(p=>p[0])),maxX=Math.max(...corners.map(p=>p[0])),minY=Math.min(...corners.map(p=>p[1])),maxY=Math.max(...corners.map(p=>p[1]));if(maxX-minX<=0||maxY-minY<=0)continue;
  const [x1,y1,x2,y2]=rect,sx=(Math.abs(x2-x1))/(maxX-minX),sy=(Math.abs(y2-y1))/(maxY-minY),tx=Math.min(x1,x2)-minX*sx,ty=Math.min(y1,y2)-minY*sy,name=page.node.newXObject('MyDocAnnot',appearance);
  page.pushOperators(pushGraphicsState(),concatTransformationMatrix(sx,0,0,sy,tx,ty),drawObject(name),popGraphicsState());
 }});
}
