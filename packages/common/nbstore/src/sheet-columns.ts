import * as Y from 'yjs';
import {workbookSheets} from './office-workbook';
export function sheetRowIdentity(row:Y.Array<string>){const item=row._item;return item?`${item.id.client}-${item.id.clock}`:'';}
export function sheetColumnIds(sheet:Y.Array<Y.Array<string>>){
 const doc=sheet.doc;if(!doc)return [];
 const id=workbookSheets(doc).find(entry=>entry.sheet===sheet)?.id;if(!id)throw new Error('工作表已删除');
 const saved=doc.share.has('sheetColumns')?doc.getMap<Y.Array<string>>('sheetColumns').get(id):undefined;
 return saved?saved.toArray():Array.from({length:Math.max(1,...sheet.toArray().map(row=>row.length))},(_,column)=>`column-${id}-${column}`);
}
export function sheetColumnIdentity(sheet:Y.Array<Y.Array<string>>,column:number){return sheetColumnIds(sheet)[column]??'';}
export function ensureSheetColumns(doc:Y.Doc,id:string,sheet:Y.Array<Y.Array<string>>){
 const map=doc.getMap<Y.Array<string>>('sheetColumns');let columns=map.get(id);if(!columns){columns=new Y.Array<string>();columns.push(sheetColumnIds(sheet));map.set(id,columns);}return columns;
}
export function validateSheetColumns(doc:Y.Doc){
 if(!doc.share.has('sheetColumns'))return;const map=doc.getMap('sheetColumns'),sheets=workbookSheets(doc);let total=0;if(map._start||map.size>100)throw new Error('工作表列标识结构无效');
 for(const [id,value]of map){const selected=sheets.find(sheet=>sheet.id===id);if(!selected||!(value instanceof Y.Array)||value._map.size||value.length<1||value.length>16384||selected.sheet.toArray().some(row=>row.length>value.length))throw new Error('工作表列标识或宽度无效');const ids=value.toArray();total+=ids.length;if(total>200000||new Set(ids).size!==ids.length||ids.some(id=>typeof id!=='string'||! /^[a-zA-Z0-9-]{1,180}$/.test(id)))throw new Error('工作表列标识重复或超限');}
}
