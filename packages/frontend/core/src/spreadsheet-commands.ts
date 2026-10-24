import type * as Y from 'yjs';
import {cellRevision,type SheetCellEdit} from '../../../common/nbstore/src/spreadsheet-operation';
import type {SpreadsheetCellType} from '../../../common/nbstore/src/spreadsheet-cell-type';
/** Pure helpers behind the spreadsheet ribbon. They only plan `setCells` edits; the server still checks every base. */
export type SheetRange={top:number;left:number;bottom:number;right:number};
export const columnName=(index:number)=>{let name='';for(let n=index+1;n;n=Math.floor((n-1)/26))name=String.fromCharCode(65+(n-1)%26)+name;return name;};
export const cellName=(row:number,column:number)=>columnName(column)+(row+1);
export function parseCellRange(text:string,rows:number,columns:number):SheetRange|undefined{
 const match=/^\s*([A-Za-z]{1,3})(\d{1,6})(?::([A-Za-z]{1,3})(\d{1,6}))?\s*$/.exec(text);if(!match)return;
 const column=(letters:string)=>[...letters.toUpperCase()].reduce((n,char)=>n*26+char.charCodeAt(0)-64,0)-1,r1=Number(match[2])-1,c1=column(match[1]),r2=match[4]?Number(match[4])-1:r1,c2=match[3]?column(match[3]):c1;
 if([r1,r2].some(r=>r<0||r>=rows)||[c1,c2].some(c=>c<0||c>=columns))return;
 return {top:Math.min(r1,r2),left:Math.min(c1,c2),bottom:Math.max(r1,r2),right:Math.max(c1,c2)};
}
const NUMBER=/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
/** Status bar figures like Excel: count of non-empty cells, and sum/average over numeric ones. */
export function selectionStats(values:string[]){
 let count=0,numeric=0,sum=0;for(const value of values){const text=value.trim();if(!text)continue;count++;if(NUMBER.test(text)){numeric++;sum+=Number(text);}}
 return {count,numeric,sum,average:numeric?sum/numeric:undefined};
}
export const formatStat=(value:number)=>Number.isInteger(value)?String(value):String(Number(value.toPrecision(12)));
/**
 * AutoSum target and range. A single cell sums the contiguous numbers directly above (or, failing that, to the left);
 * a multi-cell selection writes one formula per column in the row below it.
 */
export function autoSumPlan(grid:string[][],range:SheetRange,fn='SUM'):{row:number;column:number;formula:string}[]{
 const numeric=(r:number,c:number)=>NUMBER.test((grid[r]?.[c]??'').trim());
 if(range.top===range.bottom&&range.left===range.right){
  const {top:r,left:c}=range;let start=r;while(start>0&&numeric(start-1,c))start--;
  if(start<r)return [{row:r,column:c,formula:`=${fn}(${cellName(start,c)}:${cellName(r-1,c)})`}];
  let left=c;while(left>0&&numeric(r,left-1))left--;
  return left<c?[{row:r,column:c,formula:`=${fn}(${cellName(r,left)}:${cellName(r,c-1)})`}]:[];
 }
 const row=range.bottom+1;if(row>=grid.length)return [];
 return Array.from({length:range.right-range.left+1},(_,i)=>range.left+i).map(c=>({row,column:c,formula:`=${fn}(${cellName(range.top,c)}:${cellName(range.bottom,c)})`}));
}
/** Sort whole rows of the range by one column. Numbers before text, empty cells last, stable. */
export function sortedRows(grid:string[][],range:SheetRange,key:number,descending:boolean){return sortedRowsBy(grid,range,[{column:key,descending}]);}
export type SheetSortKey={column:number;descending:boolean};
/**
 * 自定义排序: whole rows ordered by several keys in turn (the next key breaks ties). Within a key numbers come before
 * text and empty cells last in either direction (only the order inside each group flips). `header` keeps the first row.
 */
export function sortedRowsBy(grid:string[][],range:SheetRange,keys:SheetSortKey[],header=false){
 const rows=Array.from({length:range.bottom-range.top+1},(_,i)=>range.top+i),fixed=header?rows.splice(0,1):[];
 const rank=(value:string)=>{const text=value.trim();return !text?2:NUMBER.test(text)?0:1;};
 const compare=(a:number,b:number,{column,descending}:SheetSortKey)=>{const x=grid[a]?.[column]??'',y=grid[b]?.[column]??'',rx=rank(x),ry=rank(y);if(rx!==ry)return rx-ry;if(rx===2)return 0;
  const order=rx===0?Number(x)-Number(y):x.localeCompare(y,'zh-CN',{numeric:true});return descending?-order:order;};
 return [...fixed,...rows.sort((a,b)=>{for(const key of keys){const order=compare(a,b,key);if(order)return order;}return 0;})];
}
/** Shift the number of decimals in a simple number format (`0`, `0.00`, `#,##0.0`, `0.0%`). */
export function adjustDecimals(code:string|undefined,delta:1|-1){
 const base=!code||code==='General'?'0':code,match=/^(.*?[#0])(\.(0+))?(%?)$/.exec(base);if(!match)return;
 const places=Math.max(0,Math.min(10,(match[3]?.length??0)+delta));return match[1]+(places?'.'+'0'.repeat(places):'')+match[4];
}
export function findMatches(grid:string[][],query:string,caseSensitive=false){
 if(!query)return [] as [number,number][];const needle=caseSensitive?query:query.toLowerCase(),hits:[number,number][]=[];
 grid.forEach((row,r)=>row.forEach((value,c)=>{if((caseSensitive?value:value.toLowerCase()).includes(needle))hits.push([r,c]);}));return hits;
}
/** Literal, case-sensitive replacement; formulas are left alone so references cannot be rewritten by accident. */
export function replacePlan(grid:string[][],query:string,replacement:string,skip:(r:number,c:number)=>boolean){
 if(!query)return [];const edits:{row:number;column:number;value:string}[]=[];
 grid.forEach((row,r)=>row.forEach((value,c)=>{if(value.includes(query)&&!value.trim().startsWith('=')&&!skip(r,c))edits.push({row:r,column:c,value:value.split(query).join(replacement)});}));return edits;
}
export function rangeToTsv(grid:string[][],range:SheetRange){
 const quote=(value:string)=>/[\t\n"]/.test(value)?'"'+value.replaceAll('"','""')+'"':value;
 return Array.from({length:range.bottom-range.top+1},(_,i)=>Array.from({length:range.right-range.left+1},(_,j)=>quote(grid[range.top+i]?.[range.left+j]??'')).join('\t')).join('\n');
}
export function cellEdits(sheet:Y.Array<Y.Array<string>>,edits:{row:number;column:number;value:string;cellType?:SpreadsheetCellType}[]):SheetCellEdit[]{
 return edits.map(edit=>({row:edit.row,column:edit.column,value:edit.value,base:cellRevision(sheet,edit.row,edit.column),...(edit.cellType?{cellType:edit.cellType}:{})}));
}
/** Border presets of the ribbon (所有框线、外侧框线、上/下/左/右框线、粗外框、无框线) as per-cell patches. */
export type BorderPreset='all'|'outside'|'thickOutside'|'top'|'bottom'|'left'|'right'|'none';
export function borderPatches(range:SheetRange,preset:BorderPreset,line='thin'):{row:number;column:number;format:Record<string,string|null>}[]{
 const result:{row:number;column:number;format:Record<string,string|null>}[]=[],outer=preset==='thickOutside'?'thick':line;
 for(let r=range.top;r<=range.bottom;r++)for(let c=range.left;c<=range.right;c++){
  const edge={borderTop:r===range.top,borderBottom:r===range.bottom,borderLeft:c===range.left,borderRight:c===range.right},format:Record<string,string|null>={};
  if(preset==='none')Object.assign(format,{borderTop:null,borderBottom:null,borderLeft:null,borderRight:null});
  else if(preset==='all')Object.assign(format,{borderTop:line,borderBottom:line,borderLeft:line,borderRight:line});
  else if(preset==='outside'||preset==='thickOutside'){for(const [side,on]of Object.entries(edge))if(on)format[side]=outer;}
  else{const side=({top:'borderTop',bottom:'borderBottom',left:'borderLeft',right:'borderRight'} as const)[preset];if(edge[side])format[side]=line;}
  if(Object.keys(format).length)result.push({row:r,column:c,format});
 }
 return result;
}
/**
 * 删除重复项: rows of the range whose values in `keys` (absolute columns) repeat an earlier row are removed; the kept
 * rows move up in order and the freed rows at the bottom are cleared. Text is compared case-insensitively after trimming.
 */
export function dedupePlan(grid:string[][],range:SheetRange,keys:number[],header:boolean){
 const start=range.top+(header?1:0),seen=new Set<string>(),kept:number[]=[];
 for(let r=start;r<=range.bottom;r++){const key=JSON.stringify(keys.map(c=>(grid[r]?.[c]??'').trim().toLowerCase()));if(seen.has(key))continue;seen.add(key);kept.push(r);}
 const order=[...kept,...Array.from({length:range.bottom-start+1-kept.length},()=>-1)],edits:{row:number;column:number;value:string;source:number}[]=[];
 order.forEach((source,i)=>{const target=start+i;for(let c=range.left;c<=range.right;c++){const value=source<0?'':grid[source]?.[c]??'';if(value!==(grid[target]?.[c]??''))edits.push({row:target,column:c,value,source});}});
 return {removed:range.bottom-start+1-kept.length,kept:kept.length,edits};
}
/** 分列: split one column by a delimiter into it and the columns to its right. */
export function splitPlan(grid:string[][],range:SheetRange,column:number,delimiter:string,width:number){
 if(!delimiter)return {edits:[],overwrite:0,overflow:false,columns:0};const edits:{row:number;column:number;value:string}[]=[];let overwrite=0,columns=1,overflow=false;
 for(let r=range.top;r<=range.bottom;r++){
  const value=grid[r]?.[column]??'';if(!value||value.trim().startsWith('='))continue;const parts=(delimiter===' '?value.trim().split(/\s+/):value.split(delimiter)).map(part=>part.trim());if(parts.length<2)continue;
  columns=Math.max(columns,parts.length);if(column+parts.length>width){overflow=true;continue;}
  parts.forEach((part,i)=>{const c=column+i,current=grid[r]?.[c]??'';if(i&&current!=='')overwrite++;if(part!==current)edits.push({row:r,column:c,value:part});});
 }
 return {edits,overwrite,overflow,columns};
}
