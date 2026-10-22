import * as Y from 'yjs';
export interface PdfAnnotation {page:number;type:'comment'|'text'|'highlight'|'ink';x:number;y:number;content:string;width?:number;height?:number;points?:number[];}
export type PdfOperation={operationId:string;type:'create'|'update'|'delete';annotationId:string;base:string;annotation?:PdfAnnotation};
export function pdfAnnotationId(value:Y.Map<unknown>){const item=value._item;return typeof value.get('id')==='string'?value.get('id') as string:item?`legacy-${item.id.client}-${item.id.clock}`:'';}
export function pdfAnnotationRevision(value:Y.Map<unknown>){return JSON.stringify([...value.entries()].sort(([a],[b])=>a.localeCompare(b)));}
export function validPdfAnnotation(value:unknown,strict=false):value is PdfAnnotation{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const a=value as Record<string,unknown>;
 return (!strict||Object.keys(a).every(key=>['page','type','x','y','content','width','height','points'].includes(key)))&&Number.isSafeInteger(a.page)&&Number(a.page)>=1&&Number(a.page)<=1000000&&['comment','text','highlight','ink'].includes(String(a.type))&&['x','y'].every(key=>typeof a[key]==='number'&&Number.isFinite(a[key])&&Math.abs(Number(a[key]))<=1000000)&&typeof a.content==='string'&&a.content.length<=20000&&['width','height'].every(key=>a[key]===undefined||typeof a[key]==='number'&&Number.isFinite(a[key])&&Number(a[key])>0&&Number(a[key])<=1000000)&&(a.points===undefined||Array.isArray(a.points)&&a.points.length>=4&&a.points.length<=2000&&a.points.length%2===0&&a.points.every(n=>typeof n==='number'&&Number.isFinite(n)&&Math.abs(n)<=1000000))&&(!strict||a.type!=='ink'||a.points!==undefined)&&(!strict||a.type!=='highlight'||Number(a.width)>0&&Number(a.height)>0);
}
