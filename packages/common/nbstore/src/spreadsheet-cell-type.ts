import {sheetColumnIds} from './sheet-columns';
import {workbookSheets,workbookSheet,FIRST_WORKSHEET} from './office-workbook';
import * as Y from 'yjs';
import {cellDraftKey} from './spreadsheet-operation';
export type SpreadsheetCellType='auto'|'text'|'formula';
export const validSpreadsheetCellType=(value:unknown):value is SpreadsheetCellType=>value==='auto'||value==='text'||value==='formula';
export function spreadsheetCellType(doc:Y.Doc,row:number,column:number,sheetId=FIRST_WORKSHEET):SpreadsheetCellType{
 const selected=workbookSheet(doc,sheetId);if(!doc.share.has('sheetCellTypes'))return selected.defaultType;const types=doc.getMap<SpreadsheetCellType>('sheetCellTypes'),sheet=selected.sheet;
 const legacy=JSON.stringify([Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(sheet,row))),column]);return types.get(cellDraftKey(sheet,row,column))??types.get(legacy)??selected.defaultType;
}
export function setSpreadsheetCellType(doc:Y.Doc,row:number,column:number,type:SpreadsheetCellType,sheetId=FIRST_WORKSHEET){
 if(!validSpreadsheetCellType(type))throw new Error('单元格内容类型无效');const types=doc.getMap<SpreadsheetCellType>('sheetCellTypes');if(!types.has('$default'))types.set('$default','auto');
 const selected=workbookSheet(doc,sheetId),key=cellDraftKey(selected.sheet,row,column);const legacy=JSON.stringify([Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(selected.sheet,row))),column]);types.delete(legacy);if(type===selected.defaultType)types.delete(key);else types.set(key,type);
}
function location(doc:Y.Doc,key:string,sheets=workbookSheets(doc)){
 try{const [encoded,columnKey]=JSON.parse(key);if(!Array.isArray(encoded)||encoded.length>256||encoded.some(n=>!Number.isInteger(n)||n<0||n>255))return undefined;const position=Y.createAbsolutePositionFromRelativePosition(Y.decodeRelativePosition(Uint8Array.from(encoded)),doc),selected=sheets.find(sheet=>sheet.sheet===position?.type);if(!position||!selected||position.index>=selected.sheet.length)return undefined;
 const expected=JSON.stringify([Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(selected.sheet,position.index))),columnKey]);if(expected!==key)return undefined;
 const column=typeof columnKey==='string'?sheetColumnIds(selected.sheet).indexOf(columnKey):columnKey;if(!Number.isSafeInteger(column)||column<0||column>=16384||column>=selected.sheet.get(position.index).length)return undefined;return {row:position.index,column,sheet:selected.sheet,id:selected.id};
 }catch{return undefined;}
}
export function clearRowCellTypes(doc:Y.Doc,row:number,sheetId=FIRST_WORKSHEET){if(!doc.share.has('sheetCellTypes'))return;const types=doc.getMap('sheetCellTypes'),sheets=workbookSheets(doc);for(const key of types.keys()){if(key==='$default')continue;const cell=location(doc,key,sheets);if(cell?.id===sheetId&&cell.row===row)types.delete(key);}}
export function clearSheetCellTypes(doc:Y.Doc,sheetId:string){if(!doc.share.has('sheetCellTypes'))return;const types=doc.getMap('sheetCellTypes'),sheets=workbookSheets(doc);for(const key of types.keys())if(key!=='$default'&&location(doc,key,sheets)?.id===sheetId)types.delete(key);}
export function validateSpreadsheetCellTypes(doc:Y.Doc){
 if(!doc.share.has('sheetCellTypes'))return;const types=doc.getMap('sheetCellTypes');if(!['auto','text'].includes(String(types.get('$default')))||types.size>10001)throw new Error('单元格内容类型设置无效');
 const sheets=workbookSheets(doc);for(const [key,value]of types){if(!validSpreadsheetCellType(value))throw new Error('单元格内容类型无效');if(key==='$default')continue;const cell=location(doc,key,sheets);if(!cell)throw new Error('单元格内容类型引用无效');if(value==='formula'&&!cell.sheet.get(cell.row).get(cell.column)!.trim().startsWith('='))throw new Error('公式类型的内容必须以 = 开头');}
}

export function snapshotSheetCellTypes(doc:Y.Doc,sheetId:string){if(!doc.share.has('sheetCellTypes'))return [];const sheets=workbookSheets(doc),result:{row:number;column:number;type:SpreadsheetCellType}[]=[];for(const [key,type]of doc.getMap<SpreadsheetCellType>('sheetCellTypes')){if(key==='$default')continue;const cell=location(doc,key,sheets);if(cell?.id===sheetId)result.push({row:cell.row,column:cell.column,type});}return result;}
