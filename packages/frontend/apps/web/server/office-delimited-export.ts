import {workbookSheet,workbookFormulaSnapshot,FIRST_WORKSHEET} from '../../../../common/nbstore/src/office-workbook';
import {calculateFormulaWorkbookTyped} from '../../../../common/nbstore/src/spreadsheet-formula';
import {spreadsheetCellType} from '../../../../common/nbstore/src/spreadsheet-cell-type';
import * as Y from 'yjs';
import {encodeDelimitedText,type DelimitedFormat,type DelimitedTextMode} from '../../../../common/nbstore/src/office-delimited-export';
import {validateSpreadsheetUpdate} from './office-spreadsheet';
export function exportDelimitedText(snapshot:Uint8Array,format:DelimitedFormat,mode:DelimitedTextMode,content:'text'|'values'='text',sheetId=FIRST_WORKSHEET){
 validateSpreadsheetUpdate(null,snapshot);const doc=new Y.Doc();try{
  Y.applyUpdate(doc,snapshot);if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('表格状态缺少依赖');
  if(content!=='text'&&content!=='values')throw new Error('表格导出内容设置无效');
  const selected=workbookSheet(doc,sheetId),{sheets,book}=workbookFormulaSnapshot(doc),byName=new Map(sheets.map(sheet=>[sheet.name,sheet]));
  const calculated=content==='values'?calculateFormulaWorkbookTyped(book,(name,row,column)=>spreadsheetCellType(doc,row,column,byName.get(name)!.id),[selected.name])[selected.name]:undefined;
  const rows=calculated?calculated.map(row=>row.map(value=>typeof value==='boolean'?value?'TRUE':'FALSE':String(value??0))):book[selected.name];
  return encodeDelimitedText(rows,format,mode);
 }finally{doc.destroy();}
}
