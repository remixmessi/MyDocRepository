import * as Y from 'yjs';
import {sheetColumnIds,sheetRowIdentity} from './sheet-columns';
export interface SheetCellFormat {bold?:true;italic?:true;underline?:true;color?:string;fill?:string;align?:'left'|'center'|'right';numberFormat?:string;}
export const SHEET_FORMAT_KEYS=['bold','italic','underline','color','fill','align','numberFormat'] as const;
export const SHEET_FORMAT_LIMITS={formats:50000,merges:10000,batch:2000,span:1000};
export const SHEET_NUMBER_FORMATS=['0','0.00','#,##0','#,##0.00','0%','0.00%','¥#,##0.00','$#,##0.00','yyyy-mm-dd','yyyy/m/d','yyyy年m月d日','hh:mm','hh:mm:ss','yyyy-mm-dd hh:mm','@'];
export const validSheetNumberFormat=(value:unknown):value is string=>typeof value==='string'&&value!=='General'&&value.length>=1&&value.length<=64&&!/[\x00-\x1f]/.test(value);
const color=(value:unknown)=>typeof value==='string'&&/^[0-9A-F]{6}$/.test(value);
export function validSheetCellFormat(value:unknown):value is SheetCellFormat{
 if(!value||typeof value!=='object'||Array.isArray(value)||value instanceof Y.AbstractType)return false;const v=value as Record<string,unknown>,keys=Object.keys(v);
 return keys.length>0&&keys.every(key=>(SHEET_FORMAT_KEYS as readonly string[]).includes(key))&&['bold','italic','underline'].every(key=>v[key]===undefined||v[key]===true)&&['color','fill'].every(key=>v[key]===undefined||color(v[key]))&&(v.align===undefined||['left','center','right'].includes(v.align as string))&&(v.numberFormat===undefined||validSheetNumberFormat(v.numberFormat));
}
export const validSheetMerge=(value:unknown):value is {rowSpan:number;colSpan:number}=>!!value&&typeof value==='object'&&!Array.isArray(value)&&!(value instanceof Y.AbstractType)&&Object.keys(value).length===2&&['rowSpan','colSpan'].every(key=>Number.isInteger((value as Record<string,unknown>)[key])&&Number((value as Record<string,unknown>)[key])>=1&&Number((value as Record<string,unknown>)[key])<=SHEET_FORMAT_LIMITS.span)&&(value as {rowSpan:number}).rowSpan*(value as {colSpan:number}).colSpan>=2;
/** Keys use stable row and column identities so formats and merges follow structure edits. */
export const sheetCellKey=(sheetId:string,rowKey:string,columnKey:string)=>sheetId+':'+rowKey+':'+columnKey;
export function parseSheetCellKey(key:string){const parts=key.split(':');return parts.length===3&&/^[a-zA-Z0-9-]{1,128}$/.test(parts[0])&&/^\d+-\d+$/.test(parts[1])&&/^[a-zA-Z0-9-]{1,180}$/.test(parts[2])?{sheet:parts[0],row:parts[1],column:parts[2]}:undefined;}
export function sheetCellFormat(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>,row:number,column:number):SheetCellFormat|undefined{
 if(!doc.share.has('sheetFormats'))return;const target=sheet.get(row),columnKey=sheetColumnIds(sheet)[column];if(!target||!columnKey)return;const value=doc.getMap('sheetFormats').get(sheetCellKey(sheetId,sheetRowIdentity(target),columnKey));return validSheetCellFormat(value)?value:undefined;
}
export type SheetMergeLayout={anchors:Map<string,{row:number;column:number;rowSpan:number;colSpan:number;key:string}>;owner:Map<string,string>};
/**
 * Merges resolve against the live grid. A merge whose area would cover text or overlap an earlier merge is ignored,
 * so stale or concurrent merges never hide content.
 */
export function resolveSheetMerges(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>):SheetMergeLayout{
 const anchors:SheetMergeLayout['anchors']=new Map(),owner=new Map<string,string>();if(!doc.share.has('sheetMerges'))return {anchors,owner};
 const rows=sheet.toArray(),rowIndex=new Map(rows.map((row,i)=>[sheetRowIdentity(row),i])),columns=sheetColumnIds(sheet),columnIndex=new Map(columns.map((id,i)=>[id,i]));
 const entries=[...doc.getMap('sheetMerges')].flatMap(([key,value])=>{const parsed=parseSheetCellKey(key);if(!parsed||parsed.sheet!==sheetId||!validSheetMerge(value))return [];const row=rowIndex.get(parsed.row),column=columnIndex.get(parsed.column);return row===undefined||column===undefined?[]:[{key,row,column,rowSpan:Math.min(value.rowSpan,rows.length-row),colSpan:Math.min(value.colSpan,columns.length-column)}];}).sort((a,b)=>a.row-b.row||a.column-b.column);
 for(const merge of entries){
  if(merge.rowSpan*merge.colSpan<2||owner.has(merge.row+':'+merge.column))continue;let valid=true;
  for(let r=merge.row;r<merge.row+merge.rowSpan&&valid;r++)for(let c=merge.column;c<merge.column+merge.colSpan;c++)if(owner.has(r+':'+c)||(r!==merge.row||c!==merge.column)&&(rows[r].get(c)??'')!==''){valid=false;break;}
  if(!valid)continue;for(let r=merge.row;r<merge.row+merge.rowSpan;r++)for(let c=merge.column;c<merge.column+merge.colSpan;c++)owner.set(r+':'+c,merge.row+':'+merge.column);anchors.set(merge.row+':'+merge.column,merge);
 }
 return {anchors,owner};
}
/** Keep merges and formats consistent with a row/column insert or delete. Call before mutating the grid. */
export function adjustSheetFormatsForStructure(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>,axis:'row'|'column',index:number,deleting:boolean){
 const layout=resolveSheetMerges(doc,sheetId,sheet),merges=doc.share.has('sheetMerges')?doc.getMap<{rowSpan:number;colSpan:number}>('sheetMerges'):undefined,rows=sheet.toArray(),columns=sheetColumnIds(sheet);
 for(const merge of layout.anchors.values()){
  const start=axis==='row'?merge.row:merge.column,size=axis==='row'?merge.rowSpan:merge.colSpan,span={rowSpan:merge.rowSpan,colSpan:merge.colSpan},key=axis==='row'?'rowSpan':'colSpan';
  if(!deleting){if(index>start&&index<start+size)merges!.set(merge.key,{...span,[key]:size+1});continue;}
  if(index<start||index>=start+size)continue;
  const next={...span,[key]:size-1};
  if(index>start){if(next.rowSpan*next.colSpan>=2)merges!.set(merge.key,next);else merges!.delete(merge.key);continue;}
  // Deleting the anchor's row/column hands the merge to the next row/column inside it.
  merges!.delete(merge.key);if(next.rowSpan*next.colSpan<2)continue;
  const row=axis==='row'?rows[merge.row+1]:rows[merge.row],column=axis==='row'?columns[merge.column]:columns[merge.column+1];merges!.set(sheetCellKey(sheetId,sheetRowIdentity(row),column),next);
 }
 if(!deleting)return;const removed=axis==='row'?sheetRowIdentity(rows[index]):columns[index];
 for(const root of ['sheetFormats','sheetMerges'])if(doc.share.has(root)){const map=doc.getMap(root);for(const key of [...map.keys()]){const parsed=parseSheetCellKey(key);if(parsed?.sheet===sheetId&&(axis==='row'?parsed.row:parsed.column)===removed)map.delete(key);}}
}
export function clearSheetFormats(doc:Y.Doc,sheetId:string){for(const root of ['sheetFormats','sheetMerges'])if(doc.share.has(root)){const map=doc.getMap(root);for(const key of [...map.keys()])if(parseSheetCellKey(key)?.sheet===sheetId)map.delete(key);}}

// Excel 1900 date system, including its fictitious 1900-02-29 (serial 60).
const DAY=86400000,EPOCH=Date.UTC(1899,11,30);
export function excelSerialParts(serial:number){
 if(!Number.isFinite(serial)||serial<0||serial>=2958466)return;
 let days=Math.floor(serial),time=Math.round((serial-days)*86400);if(time===86400){days++;time=0;}let date:Date;
 if(days===60)return {year:1900,month:2,day:29,hour:Math.floor(time/3600)%24,minute:Math.floor(time/60)%60,second:time%60};
 date=new Date(EPOCH+(days<60?days+1:days)*DAY);
 return {year:date.getUTCFullYear(),month:date.getUTCMonth()+1,day:date.getUTCDate(),hour:Math.floor(time/3600)%24,minute:Math.floor(time/60)%60,second:time%60};
}
export function excelSerialFromParts(year:number,month:number,day:number,hour=0,minute=0,second=0){
 const utc=Date.UTC(year,month-1,day),date=new Date(utc);if(year<1900||year>9999||date.getUTCFullYear()!==year||date.getUTCMonth()!==month-1||date.getUTCDate()!==day||hour>23||minute>59||second>59)return;
 let days=Math.round((utc-EPOCH)/DAY);if(days<=60)days--;return days+(hour*3600+minute*60+second)/86400;
}
const unquoted=(code:string)=>code.replace(/"[^"]*"|\\./g,'').replace(/\[[^\]]*\]/g,'');
export const isDateNumberFormat=(code:string|undefined)=>!!code&&/[ymdhs]/i.test(unquoted(code.split(';')[0]));
const pad=(value:number,size=2)=>String(value).padStart(size,'0');
function formatDate(serial:number,code:string){
 const parts=excelSerialParts(serial);if(!parts)return;const tokens=code.split(';')[0].match(/"[^"]*"|\\.|\[[^\]]*\]|AM\/PM|y+|m+|d+|h+|s+|[^ymdhsA"\\[]+|A/gi)??[];let output='';
 const hour12=tokens.some(token=>/^AM\/PM$/i.test(token));
 tokens.forEach((token,i)=>{
  const lower=token.toLowerCase(),prev=tokens.slice(0,i).reverse().find(t=>/^[ymdhs]/i.test(t))?.toLowerCase(),next=tokens.slice(i+1).find(t=>/^[ymdhs]/i.test(t))?.toLowerCase();
  if(token.startsWith('"'))output+=token.slice(1,-1);else if(token.startsWith('\\'))output+=token.slice(1);else if(token.startsWith('['))return;
  else if(/^y+$/.test(lower))output+=lower.length<=2?pad(parts.year%100):String(parts.year);
  else if(/^m+$/.test(lower)){const minute=prev?.startsWith('h')||next?.startsWith('s');output+=minute?(lower.length>1?pad(parts.minute):String(parts.minute)):lower.length>1?pad(parts.month):String(parts.month);}
  else if(/^d+$/.test(lower))output+=lower.length>1?pad(parts.day):String(parts.day);
  else if(/^h+$/.test(lower)){const hour=hour12?(parts.hour%12||12):parts.hour;output+=lower.length>1?pad(hour):String(hour);}
  else if(/^s+$/.test(lower))output+=lower.length>1?pad(parts.second):String(parts.second);
  else if(/^AM\/PM$/i.test(token))output+=parts.hour<12?'AM':'PM';else output+=token;
 });
 return output;
}
function formatNumber(value:number,code:string){
 const sections=code.split(';'),section=value<0&&sections[1]!==undefined?sections[1]:sections[0],negative=value<0&&sections[1]===undefined;
 const literal=section.replace(/\[[^\]]*\]/g,'').replace(/"([^"]*)"/g,'$1').replace(/\\(.)/g,'$1'),core=/[#0][#0,.]*%?/.exec(literal)?.[0];if(!core)return;
 const percent=core.endsWith('%'),decimals=/\.(0+)/.exec(core)?.[1].length??0,grouping=core.includes(','),amount=Math.abs(value)*(percent?100:1);
 let text=amount.toFixed(decimals);if(grouping){const [int,frac]=text.split('.');text=int.replace(/\B(?=(\d{3})+(?!\d))/g,',')+(frac?'.'+frac:'');}
 const index=literal.indexOf(core);return (negative?'-':'')+literal.slice(0,index)+text+(percent?'%':'')+literal.slice(index+core.length);
}
/** Display text for a stored value; non-numeric values and unsupported codes show as stored. */
export function formatSheetValue(value:string,code:string|undefined){
 if(!code||code==='General'||code==='@')return value;const trimmed=value.trim();if(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(trimmed))return value;const number=Number(trimmed);
 try{return (isDateNumberFormat(code)?formatDate(number,code):formatNumber(number,code))??value;}catch{return value;}
}
/** Editable text for date-formatted cells, and the serial to store when the user types a date. */
export function sheetDateEditText(value:string){const n=Number(value.trim());const parts=value.trim()&&Number.isFinite(n)?excelSerialParts(n):undefined;if(!parts)return value;const date=parts.year+'-'+pad(parts.month)+'-'+pad(parts.day);return n%1?date+' '+pad(parts.hour)+':'+pad(parts.minute)+':'+pad(parts.second):date;}
export function parseSheetDateInput(value:string){
 const match=/^\s*(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$/.exec(value);if(!match)return;
 const serial=excelSerialFromParts(Number(match[1]),Number(match[2]),Number(match[3]),Number(match[4]??0),Number(match[5]??0),Number(match[6]??0));return serial===undefined?undefined:String(Math.round(serial*1e10)/1e10);
}
