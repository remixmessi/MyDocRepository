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
export function sortedRows(grid:string[][],range:SheetRange,key:number,descending:boolean){
 const rows=Array.from({length:range.bottom-range.top+1},(_,i)=>range.top+i);
 const rank=(value:string)=>{const text=value.trim();return !text?2:NUMBER.test(text)?0:1;};
 return [...rows].sort((a,b)=>{const x=grid[a]?.[key]??'',y=grid[b]?.[key]??'',rx=rank(x),ry=rank(y);if(rx!==ry)return rx-ry;if(rx===2)return 0;
  const order=rx===0?Number(x)-Number(y):x.localeCompare(y,'zh-CN',{numeric:true});return descending?-order:order;});
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
