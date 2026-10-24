import * as Y from 'yjs';
/** `strokes` (ink only): point-pair counts of each stroke in `points`, so a signature keeps its separate strokes. */
export interface PdfAnnotation {page:number;type:'comment'|'text'|'highlight'|'ink';x:number;y:number;content:string;width?:number;height?:number;points?:number[];strokes?:number[];}
/**
 * Page arrangement (页面管理): the original PDF never changes; the order, rotation and deletion of its pages and any
 * inserted blank pages are a list of entries. `source` is a 1-based page of the original. Annotations keep original
 * page numbers and travel with their page.
 */
export type PdfPageEntry={id:string;source:number;rotate?:number}|{id:string;blank:{width:number;height:number};rotate?:number};
export type PdfWatermark={text:string;color:string;opacity:number;size:number;rotation:number};
export type PdfAnnotationOperation={operationId:string;type:'create'|'update'|'delete';annotationId:string;base:string;annotation?:PdfAnnotation};
export type PdfOperation=PdfAnnotationOperation|{operationId:string;type:'pages';base:string;pages:PdfPageEntry[]}|{operationId:string;type:'watermark';base:string;watermark:PdfWatermark|null};
export const PDF_PAGE_LIMIT=10000;
export function pdfAnnotationId(value:Y.Map<unknown>){const item=value._item;return typeof value.get('id')==='string'?value.get('id') as string:item?`legacy-${item.id.client}-${item.id.clock}`:'';}
export function pdfAnnotationRevision(value:Y.Map<unknown>){return JSON.stringify([...value.entries()].sort(([a],[b])=>a.localeCompare(b)));}
export function validPdfAnnotation(value:unknown,strict=false):value is PdfAnnotation{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const a=value as Record<string,unknown>;
 return (!strict||Object.keys(a).every(key=>['page','type','x','y','content','width','height','points','strokes'].includes(key)))&&Number.isSafeInteger(a.page)&&Number(a.page)>=1&&Number(a.page)<=1000000&&['comment','text','highlight','ink'].includes(String(a.type))&&['x','y'].every(key=>typeof a[key]==='number'&&Number.isFinite(a[key])&&Math.abs(Number(a[key]))<=1000000)&&typeof a.content==='string'&&a.content.length<=20000&&['width','height'].every(key=>a[key]===undefined||typeof a[key]==='number'&&Number.isFinite(a[key])&&Number(a[key])>0&&Number(a[key])<=1000000)&&(a.points===undefined||Array.isArray(a.points)&&a.points.length>=4&&a.points.length<=2000&&a.points.length%2===0&&a.points.every(n=>typeof n==='number'&&Number.isFinite(n)&&Math.abs(n)<=1000000))&&(!strict||a.type!=='ink'||a.points!==undefined)&&(!strict||a.type!=='highlight'||Number(a.width)>0&&Number(a.height)>0)
  &&(a.strokes===undefined||a.type==='ink'&&Array.isArray(a.points)&&Array.isArray(a.strokes)&&a.strokes.length>=1&&a.strokes.length<=100&&a.strokes.every(n=>Number.isSafeInteger(n)&&Number(n)>=2)&&a.strokes.reduce((sum:number,n)=>sum+Number(n),0)*2===a.points.length);
}
/** Ink paths of an annotation: one per stroke (older annotations are a single stroke). */
export function pdfInkPaths(a:Pick<PdfAnnotation,'points'|'strokes'>){
 if(!a.points)return [] as number[][];if(!a.strokes)return [a.points];const paths:number[][]=[];let at=0;for(const count of a.strokes){paths.push(a.points.slice(at,at+count*2));at+=count*2;}return paths;
}
const pageSize=(value:unknown)=>typeof value==='number'&&Number.isFinite(value)&&value>=36&&value<=14400;
export function validPdfPages(value:unknown):value is PdfPageEntry[]{
 if(!Array.isArray(value)||!value.length||value.length>PDF_PAGE_LIMIT)return false;const ids=new Set<string>();
 return value.every(entry=>{if(!entry||typeof entry!=='object'||Array.isArray(entry))return false;const e=entry as Record<string,unknown>;
  if(typeof e.id!=='string'||!/^[a-zA-Z0-9-]{1,40}$/.test(e.id)||ids.has(e.id))return false;ids.add(e.id);
  if(e.rotate!==undefined&&![90,180,270].includes(e.rotate as number))return false;
  if('blank' in e)return Object.keys(e).every(key=>['id','blank','rotate'].includes(key))&&!!e.blank&&typeof e.blank==='object'&&Object.keys(e.blank as object).length===2&&pageSize((e.blank as {width:unknown}).width)&&pageSize((e.blank as {height:unknown}).height);
  return Object.keys(e).every(key=>['id','source','rotate'].includes(key))&&Number.isSafeInteger(e.source)&&Number(e.source)>=1&&Number(e.source)<=PDF_PAGE_LIMIT;});
}
export function validPdfWatermark(value:unknown):value is PdfWatermark{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const w=value as Record<string,unknown>;
 return Object.keys(w).length===5&&typeof w.text==='string'&&w.text.trim()!==''&&w.text===w.text.trim()&&[...w.text].length<=50&&!/[\x00-\x1f\u2028\u2029]/.test(w.text)&&typeof w.color==='string'&&/^#[0-9A-F]{6}$/.test(w.color)
  &&typeof w.opacity==='number'&&w.opacity>=0.05&&w.opacity<=1&&Number.isInteger(w.size)&&Number(w.size)>=12&&Number(w.size)<=200&&[0,45,-45,90].includes(w.rotation as number);
}
export const pdfIdentityPages=(count:number):PdfPageEntry[]=>Array.from({length:Math.max(1,Math.min(count,PDF_PAGE_LIMIT))},(_,i)=>({id:'p'+(i+1),source:i+1}));
/** Stored arrangement, or undefined when the pages are still the original's (identity). */
export function pdfPages(doc:Y.Doc){const value=doc.share.has('pdfPages')?doc.getMap('pdfPages').get('arrangement'):undefined;return validPdfPages(value)?value:undefined;}
export function pdfWatermark(doc:Y.Doc){const value=doc.share.has('pdfPages')?doc.getMap('pdfPages').get('watermark'):undefined;return validPdfWatermark(value)?value:undefined;}
/** Revisions for compare-and-set: the stored JSON ('' when absent). */
export const pdfPagesRevision=(doc:Y.Doc)=>{const value=pdfPages(doc);return value?JSON.stringify(value):'';};
export const pdfWatermarkRevision=(doc:Y.Doc)=>{const value=pdfWatermark(doc);return value?JSON.stringify(value):'';};
/** "1-3,5" (1-based positions) → sorted unique positions; undefined when invalid or out of range. */
export function parsePdfPageRange(text:string,count:number){
 const result=new Set<number>();for(const part of text.split(/[,，、\s]+/).filter(Boolean)){const match=/^(\d+)(?:\s*[-–~]\s*(\d+))?$/.exec(part);if(!match)return;const a=Number(match[1]),b=Number(match[2]??match[1]);if(a<1||b<1||a>count||b>count)return;for(let n=Math.min(a,b);n<=Math.max(a,b);n++)result.add(n);}
 return result.size?[...result].sort((a,b)=>a-b):undefined;
}
