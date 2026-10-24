import {applyPdfFormValues} from './office-pdf-forms';
import {removeImportedNatives,flattenNativeAppearances} from './office-pdf-native';
import {PDFDocument,PDFName,PDFHexString,PDFString,PDFDict,PDFRef,PDFArray,PDFNumber,rgb,degrees} from 'pdf-lib';
import {PdfFonts} from './office-pdf-font';
import {validOfficeComment,type OfficeCommentThread} from '../../../../common/nbstore/src/office-comments';
import * as Y from 'yjs';
import {validatePdfUpdate} from './office-pdf';
import {pdfAnnotationId,pdfInkPaths,pdfPages,pdfWatermark,pdfIdentityPages,type PdfAnnotation,type PdfPageEntry,type PdfWatermark} from '../../../../common/nbstore/src/office-pdf';
const pdfDate=(time:number)=>{const d=new Date(time),two=(n:number)=>String(n).padStart(2,'0');return PDFString.of(`D:${d.getUTCFullYear()}${two(d.getUTCMonth()+1)}${two(d.getUTCDate())}${two(d.getUTCHours())}${two(d.getUTCMinutes())}${two(d.getUTCSeconds())}Z`);};
/**
 * File discussions become native PDF notes: the first message is a Text annotation beside the PDF annotation it
 * discusses (page 1 top-left when it has none or it was deleted), replies point to it with IRT, and a resolved
 * thread gets a hidden "Accepted" review-state reply, which is how Acrobat records a resolved note.
 */
function addDiscussionNotes(pdf:PDFDocument,doc:Y.Doc,kept:Set<number>,fallback:number){
 if(!doc.share.has('officeComments'))return;
 const annotations=new Map(doc.getArray<Y.Map<unknown>>('annotations').toArray().map(value=>[pdfAnnotationId(value),value.toJSON() as unknown as PdfAnnotation]));
 const threads=[...doc.getMap<OfficeCommentThread>('officeComments').values()].filter(validOfficeComment).sort((a,b)=>a.messages[0].created-b.messages[0].created||a.id.localeCompare(b.id));
 for(const thread of threads){
  // A note whose page was deleted or not extracted goes to the first kept page.
  const found=thread.anchor?.kind==='pdf'?annotations.get(thread.anchor.annotation):undefined,target=found&&found.page<=pdf.getPageCount()&&kept.has(found.page)?found:undefined,pageNumber=target?target.page:fallback,page=pdf.getPage(pageNumber-1),box=page.getCropBox();
  const size=20,x=Math.min(Math.max(target?target.x:box.x,box.x),box.x+box.width-size),y=Math.min(Math.max(target?target.y+(target.height??20)-size:box.y+box.height-size,box.y),box.y+box.height-size);
  let root:PDFRef|undefined;
  const note=(fields:Record<string,unknown>)=>{const ref=pdf.context.register(pdf.context.obj({Type:'Annot',Subtype:'Text',Rect:[x,y,x+size,y+size],F:4,C:[0.95,0.65,0.1],Name:'Comment',...fields}));page.node.addAnnot(ref);return ref;};
  for(const message of thread.messages){
   const ref=note({Contents:PDFHexString.fromText(message.body),T:PDFHexString.fromText(message.authorName),M:pdfDate(message.created),NM:PDFHexString.fromText('mydoc-discussion-'+message.id),...(root?{IRT:root,RT:'R'}:{Open:false})});root??=ref;
  }
  // Hidden (F=2) state replies are how viewers mark a note resolved without drawing another icon.
  if(thread.resolved)note({F:2,IRT:root,State:PDFHexString.fromText('Accepted'),StateModel:PDFHexString.fromText('Review'),Contents:PDFHexString.fromText('已解决'),T:PDFHexString.fromText(thread.messages[0].authorName),M:pdfDate(thread.updated),NM:PDFHexString.fromText('mydoc-discussion-state-'+thread.id)});
 }
}
/** A bounded binary envelope: 4-byte model length, model, immutable original. */
export function packPdfExport(state:Uint8Array,original:Uint8Array){const header=Buffer.alloc(4);header.writeUInt32LE(state.length);return Buffer.concat([header,state,original]);}
/** `font` is the optional Unicode appearance font (form values and flattened notes that Helvetica cannot encode). */
export async function exportAnnotatedPdf(packet:Uint8Array,flatten=false,font?:Uint8Array,positions?:number[]){
 const bytes=Buffer.from(packet);if(bytes.length<6)throw new Error('PDF 导出输入无效');const length=bytes.readUInt32LE(0);if(length>16*1024*1024||length+4>=bytes.length||bytes.length-length-4>20*1024*1024)throw new Error('PDF 导出输入超过限制');
 const state=bytes.subarray(4,4+length);validatePdfUpdate(null,state);let pdf:PDFDocument;try{pdf=await PDFDocument.load(bytes.subarray(4+length),{updateMetadata:false,throwOnInvalidObject:true});}catch(error){throw new Error(error instanceof Error&&error.name==='EncryptedPDFError'?'加密 PDF 暂不支持批注副本':'PDF 原件损坏或包含无法处理的结构');}if(pdf.isEncrypted)throw new Error('加密 PDF 暂不支持批注副本');if(pdf.getPageCount()>10000)throw new Error('PDF 页数超过导出上限');
 // A signed document cannot retain its signature validity after a rewrite.
 for(const [,object]of pdf.context.enumerateIndirectObjects()){if(object instanceof PDFDict){const value=object.get(PDFName.of('FT'));if(value===PDFName.of('Sig')&&object.has(PDFName.of('V'))||object.get(PDFName.of('Type'))===PDFName.of('Sig'))throw new Error('已签名 PDF 仅允许保留原件，不能生成可能误导的签名副本');}}
 if(flatten&&pdf.catalog.lookupMaybe(PDFName.of('AcroForm'),PDFDict)?.has(PDFName.of('XFA')))throw new Error('XFA 表单不支持扁平化');
 const doc=new Y.Doc();try{
  Y.applyUpdate(doc,state);
  // The arranged pages (all, or the extracted positions) decide which original pages are written and in what order.
  const original=pdf.getPageCount(),arrangement=pdfPages(doc)??pdfIdentityPages(original);
  for(const entry of arrangement)if('source' in entry&&entry.source>original)throw new Error('页面设置引用了原件中不存在的页，请重新调整页面');
  const entries=positions?positions.map(position=>{const entry=arrangement[position-1];if(!entry)throw new Error('提取的页码超出范围');return entry;}):arrangement;if(!entries.length)throw new Error('至少需要一页');
  const kept=new Set(entries.flatMap(entry=>'source' in entry?[entry.source]:[])),identity=!positions&&!pdfPages(doc),fallback=[...kept][0]??1;
  // Originals imported into the model are replaced by their (possibly edited or deleted) model copies.
  removeImportedNatives(pdf,new Set(doc.share.has('pdfSource')?[...doc.getMap('pdfSource').keys()].filter(key=>key!=='$imported'):[]));
  if(flatten)flattenNativeAppearances(pdf);
  const fonts=new PdfFonts(pdf,font);
  await applyPdfFormValues(pdf,doc,fonts);for(const annotation of doc.getArray<Y.Map<unknown>>('annotations')){
   const a=annotation.toJSON() as unknown as PdfAnnotation;if(a.page<=pdf.getPageCount()&&!kept.has(a.page))continue;if(a.type==='ink'&&!a.points||a.type==='highlight'&&(!a.width||!a.height))throw new Error('旧批注缺少几何信息，请重新定位后导出');if(a.page>pdf.getPageCount())throw new Error('批注页码超出原件范围');const page=pdf.getPage(a.page-1),box=page.getCropBox(),w=a.width??20,h=a.height??20;
   if(a.x<box.x||a.y<box.y||a.x+w>box.x+box.width||a.y+h>box.y+box.height)throw new Error('批注超出 PDF 裁剪区域，请重新定位');
   if(a.points&&a.points.some((coordinate,index)=>index%2===0?coordinate<a.x||coordinate>a.x+w:coordinate<a.y||coordinate>a.y+h))throw new Error('手绘路径超出批注区域');
   if(flatten){
    if(a.type==='highlight')page.drawRectangle({x:a.x,y:a.y,width:w,height:h,color:rgb(1,.85,0),opacity:.35});
    else if(a.type==='ink'){for(const path of pdfInkPaths(a))for(let i=2;i<path.length;i+=2)page.drawLine({start:{x:path[i-2],y:path[i-1]},end:{x:path[i],y:path[i+1]},color:rgb(.15,.35,.9),thickness:2});}
    else{const font=await fonts.for(a.content,'附注'),width=a.width??Math.min(200,box.x+box.width-a.x),height=a.height??Math.min(80,box.y+box.height-a.y),lines:string[]=[];let line='';for(const char of a.content){if(char==='\n'){lines.push(line);line='';continue;}if(line&&font.widthOfTextAtSize(line+char,10)>width){lines.push(line);line='';}line+=char;}lines.push(line);if(lines.length*12>height)throw new Error('附注内容超出扁平化区域，请缩短内容或使用普通批注副本');for(const [index,text]of lines.entries())page.drawText(text,{x:a.x,y:a.y+height-12*(index+1),size:10,font,color:rgb(.15,.35,.9)});}
    continue;
   }
   const subtype=a.type==='highlight'?'Highlight':a.type==='ink'?'Ink':'Text';
   const dict=pdf.context.obj({Type:'Annot',Subtype:subtype,Rect:[a.x,a.y,a.x+w,a.y+h],Contents:PDFHexString.fromText(a.content),NM:PDFHexString.fromText('mydoc-'+pdfAnnotationId(annotation)),T:PDFHexString.fromText(String(annotation.get('author')??'MyDoc')),F:4,C:a.type==='highlight'?[1,0.85,0]:[0.15,0.35,0.9],...(a.type==='highlight'?{QuadPoints:[a.x,a.y+h,a.x+w,a.y+h,a.x,a.y,a.x+w,a.y],CA:0.35}:{}),...(a.type==='ink'?{InkList:pdfInkPaths(a),BS:{W:2}}:{Name:'Comment'})});page.node.addAnnot(pdf.context.register(dict));
  }
  if(!flatten)addDiscussionNotes(pdf,doc,kept,fallback);
  if(flatten){
   if(pdf.catalog.has(PDFName.of('AcroForm')))pdf.getForm().flatten({updateFieldAppearances:false});
   // Everything visual is now page content; only hyperlinks stay interactive.
   for(const page of pdf.getPages()){const list=page.node.Annots();if(!list)continue;for(let index=list.size()-1;index>=0;index--){const value=list.lookup(index);if(!(value instanceof PDFDict&&value.get(PDFName.of('Subtype'))===PDFName.of('Link')))list.remove(index);}if(!list.size())page.node.delete(PDFName.of('Annots'));}
  }
  if(!identity)await arrangePages(pdf,entries,kept,original,pdfWatermark(doc),fonts);else{const watermark=pdfWatermark(doc);if(watermark)for(const page of pdf.getPages())await drawWatermark(page,watermark,page.getRotation().angle,fonts);}
  return await pdf.save({useObjectStreams:true,updateFieldAppearances:false});
 }finally{doc.destroy();}
}

type PdfPageHandle=ReturnType<PDFDocument['getPage']>;
/** Semi-transparent text centred on the page as the reader sees it (the page's own /Rotate is undone). */
async function drawWatermark(page:PdfPageHandle,watermark:PdfWatermark,pageRotation:number,fonts:PdfFonts){
 const font=await fonts.for(watermark.text,'水印'),box=page.getCropBox(),width=font.widthOfTextAtSize(watermark.text,watermark.size),height=font.heightAtSize(watermark.size,{descender:false});
 const angle=((watermark.rotation-pageRotation)%360+360)%360,radians=angle*Math.PI/180,cx=box.x+box.width/2,cy=box.y+box.height/2;
 // drawText rotates around its start point, so start half the text width back along the angle and half its height below.
 const x=cx-Math.cos(radians)*width/2+Math.sin(radians)*height/2,y=cy-Math.sin(radians)*width/2-Math.cos(radians)*height/2;
 const [r,g,b]=[1,3,5].map(i=>parseInt(watermark.color.slice(i,i+2),16)/255);
 page.drawText(watermark.text,{x,y,size:watermark.size,font,color:rgb(r,g,b),opacity:watermark.opacity,rotate:degrees(angle)});
}
/**
 * Write the arranged pages: forms are flattened when a page with fields is left out (their widgets would dangle),
 * blank pages are added, rotation and the watermark applied, and the page tree rebuilt flat in the new order.
 */
async function arrangePages(pdf:PDFDocument,entries:PdfPageEntry[],kept:Set<number>,original:number,watermark:PdfWatermark|undefined,fonts:PdfFonts){
 const pages=pdf.getPages();
 if(kept.size<original&&pdf.catalog.has(PDFName.of('AcroForm'))&&pdf.getForm().getFields().length)pdf.getForm().flatten({updateFieldAppearances:false});
 const handles=entries.map(entry=>'source' in entry?pages[entry.source-1]:pdf.addPage([entry.blank.width,entry.blank.height]));
 for(const [i,entry]of entries.entries()){const page=handles[i],angle=(page.getRotation().angle+(entry.rotate??0))%360;if(entry.rotate)page.setRotation(degrees(angle));if(watermark)await drawWatermark(page,watermark,angle,fonts);}
 const root=pdf.catalog.lookup(PDFName.of('Pages'),PDFDict),rootRef=pdf.catalog.get(PDFName.of('Pages')) as PDFRef;
 // Attributes inherited from intermediate page-tree nodes are copied onto each page before the tree is flattened.
 for(const page of handles){for(const key of ['Resources','MediaBox','CropBox','Rotate']){const name=PDFName.of(key);if(page.node.has(name))continue;let parent=page.node.lookupMaybe(PDFName.of('Parent'),PDFDict);while(parent&&!parent.has(name))parent=parent.lookupMaybe(PDFName.of('Parent'),PDFDict);const value=parent?.get(name);if(value)page.node.set(name,value);}page.node.set(PDFName.of('Parent'),rootRef);}
 const refs=handles.map(page=>page.ref);if(new Set(refs.map(String)).size!==refs.length)throw new Error('同一原件页不能在提取结果中出现两次');
 root.set(PDFName.of('Kids'),pdf.context.obj(refs) as PDFArray);root.set(PDFName.of('Count'),PDFNumber.of(refs.length));
}
