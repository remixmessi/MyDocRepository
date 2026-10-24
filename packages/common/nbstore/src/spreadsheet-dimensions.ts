import * as Y from 'yjs';
import {sheetColumnIds,sheetRowIdentity} from './sheet-columns';
/**
 * Column widths, row heights and hidden rows/columns (`sheetDimensions` root). Keys are `sheet:r:rowKey` or
 * `sheet:c:columnKey` with stable identities, so sizes follow inserts and deletes; only non-default entries are stored.
 * Sizes are CSS pixels at 100% zoom.
 */
export type SheetAxis='row'|'column';
export type SheetDimension={size?:number;hidden?:true};
export const SHEET_DIMENSION_LIMITS={entries:20000,batch:2000,column:{min:20,max:1000,default:96},row:{min:16,max:600,default:25}};
export const sheetDimensionKey=(sheet:string,axis:SheetAxis,key:string)=>sheet+':'+(axis==='row'?'r':'c')+':'+key;
export function parseSheetDimensionKey(value:string){const parts=value.split(':');return parts.length===3&&/^[a-zA-Z0-9-]{1,128}$/.test(parts[0])&&(parts[1]==='r'&&/^\d+-\d+$/.test(parts[2])||parts[1]==='c'&&/^[a-zA-Z0-9-]{1,180}$/.test(parts[2]))?{sheet:parts[0],axis:(parts[1]==='r'?'row':'column') as SheetAxis,key:parts[2]}:undefined;}
export const validSheetSize=(axis:SheetAxis,value:unknown)=>Number.isInteger(value)&&Number(value)>=SHEET_DIMENSION_LIMITS[axis].min&&Number(value)<=SHEET_DIMENSION_LIMITS[axis].max;
export function validSheetDimension(axis:SheetAxis,value:unknown):value is SheetDimension{
 if(!value||typeof value!=='object'||Array.isArray(value)||value instanceof Y.AbstractType)return false;const v=value as Record<string,unknown>,keys=Object.keys(v);
 return keys.length>0&&keys.every(key=>key==='size'||key==='hidden')&&(v.size===undefined||validSheetSize(axis,v.size))&&(v.hidden===undefined||v.hidden===true);
}
/** Dimensions of one sheet by current index. */
export function sheetDimensions(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>){
 const rows=sheet.toArray(),columns=sheetColumnIds(sheet),map=doc.share.has('sheetDimensions')?doc.getMap<SheetDimension>('sheetDimensions'):undefined;
 const read=(axis:SheetAxis,key:string)=>{const value=map?.get(sheetDimensionKey(sheetId,axis,key));return validSheetDimension(axis,value)?value:undefined;};
 return {rows:rows.map(row=>read('row',sheetRowIdentity(row))),columns:columns.map(column=>read('column',column))};
}
/** Remove the dimension entry of a deleted row/column (call before mutating the grid), or every entry of a deleted sheet. */
export function clearSheetDimensions(doc:Y.Doc,sheetId:string,removed?:{axis:SheetAxis;key:string}){
 if(!doc.share.has('sheetDimensions'))return;const map=doc.getMap('sheetDimensions');
 if(removed){map.delete(sheetDimensionKey(sheetId,removed.axis,removed.key));return;}
 for(const key of [...map.keys()])if(parseSheetDimensionKey(key)?.sheet===sheetId)map.delete(key);
}
export function validateSheetDimensions(doc:Y.Doc,sheetIds:Set<string>){
 if(!doc.share.has('sheetDimensions'))return;const map=doc.getMap('sheetDimensions');
 if(map._start||map.size>SHEET_DIMENSION_LIMITS.entries||[...map].some(([key,value])=>{const parsed=parseSheetDimensionKey(key);return !parsed||!sheetIds.has(parsed.sheet)||!validSheetDimension(parsed.axis,value);}))throw new Error('表格行高列宽或隐藏设置无效');
}
/** Excel column width (characters of the default font) ↔ pixels, and row height points ↔ pixels. */
export const excelWidthToPixels=(width:number)=>Math.round(width*7+5);
export const pixelsToExcelWidth=(pixels:number)=>Math.max(0,Math.round((pixels-5)/7*256)/256);
export const pointsToPixels=(points:number)=>Math.round(points*4/3);
export const pixelsToPoints=(pixels:number)=>Math.round(pixels*0.75*100)/100;
