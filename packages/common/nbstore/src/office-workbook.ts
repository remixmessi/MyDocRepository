import * as Y from 'yjs';
import {spreadsheetSheetName} from './office-xlsx-source';
export const FIRST_WORKSHEET='first',MAX_WORKBOOK_SHEETS=100;
export interface WorkbookSheet {id:string;name:string;sheet:Y.Array<Y.Array<string>>;defaultType:'auto'|'text';}
export function workbookSheets(doc:Y.Doc):WorkbookSheet[]{
 const source=doc.share.has('sheetXlsx')?doc.getMap('sheetXlsx').toJSON():undefined;
 const ids=doc.share.has('sheetOrder')?doc.getArray<string>('sheetOrder').toArray():[FIRST_WORKSHEET];
 return ids.map(id=>{const meta=doc.share.has('workbookSheets')?doc.getMap<Y.Map<unknown>>('workbookSheets').get(id):undefined;return {id,name:String(meta?.get('name')??(id===FIRST_WORKSHEET?spreadsheetSheetName(source):'')),sheet:id===FIRST_WORKSHEET?doc.getArray<Y.Array<string>>('sheet'):meta?.get('sheet') as Y.Array<Y.Array<string>>,defaultType:(meta?.get('defaultType')??(doc.share.has('sheetCellTypes')?doc.getMap('sheetCellTypes').get('$default'):'auto')??'auto') as 'auto'|'text'};});
}
export function workbookSheet(doc:Y.Doc,id=FIRST_WORKSHEET):WorkbookSheet{
 const exists=doc.share.has('sheetOrder')?doc.getArray<string>('sheetOrder').toArray().includes(id):id===FIRST_WORKSHEET;if(!exists)throw new Error('工作表不存在或已删除');
 const meta=doc.share.has('workbookSheets')?doc.getMap<Y.Map<unknown>>('workbookSheets').get(id):undefined;
 return {id,name:String(meta?.get('name')??spreadsheetSheetName(doc.share.has('sheetXlsx')?doc.getMap('sheetXlsx').toJSON():undefined)),sheet:id===FIRST_WORKSHEET?doc.getArray<Y.Array<string>>('sheet'):meta!.get('sheet') as Y.Array<Y.Array<string>>,defaultType:(meta?.get('defaultType')??(doc.share.has('sheetCellTypes')?doc.getMap('sheetCellTypes').get('$default'):'auto')??'auto') as 'text'|'auto'};
}
export function workbookFormulaSnapshot(doc:Y.Doc){const sheets=workbookSheets(doc),book:Record<string,string[][]>=Object.create(null);for(const sheet of sheets)book[sheet.name]=sheet.sheet.toArray().map(row=>row.toArray());return {sheets,book};}
export const validWorksheetName=(name:unknown):name is string=>typeof name==='string'&&!!name.trim()&&name.length<=31&&!/[\x00-\x1f\\/*?:\[\]]/.test(name)&&!name.startsWith("'")&&!name.endsWith("'");
