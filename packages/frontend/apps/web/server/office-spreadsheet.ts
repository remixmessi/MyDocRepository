import {validateSheetColumns} from '../../../../common/nbstore/src/sheet-columns';
import {validateOfficeComments} from '../../../../common/nbstore/src/office-comments';
import {workbookSheets,FIRST_WORKSHEET,MAX_WORKBOOK_SHEETS,validWorksheetName} from '../../../../common/nbstore/src/office-workbook';
import {validXlsxSource} from '../../../../common/nbstore/src/office-xlsx-source';
import {validateSpreadsheetCellTypes} from '../../../../common/nbstore/src/spreadsheet-cell-type';
import {prepareDelimitedImport,validDelimitedImportSettings,type DelimitedImportOptions} from '../../../../common/nbstore/src/office-delimited-import';
import * as Y from 'yjs';
import {parseDelimited} from '../../../../common/nbstore/src/office';
import {validSheetCellFormat,validSheetMerge,parseSheetCellKey,SHEET_FORMAT_LIMITS} from '../../../../common/nbstore/src/spreadsheet-format';

export const MAX_SPREADSHEET_CELLS=200_000;
export function encodeDelimitedWorkbook(text:string,delimiter:','|'\t'){
 const rows=parseDelimited(text,delimiter);if(rows.reduce((n,row)=>n+row.length,0)>MAX_SPREADSHEET_CELLS)throw new Error('表格单元格超过 20 万');
 const doc=new Y.Doc(),sheet=doc.getArray<Y.Array<string>>('sheet');
 for(const row of rows){const target=new Y.Array<string>();target.insert(0,row);sheet.push([target]);}
 return {doc,update:Y.encodeStateAsUpdate(doc)};
}
export function encodeDelimitedOriginal(bytes:Uint8Array,options:DelimitedImportOptions){
 const imported=prepareDelimitedImport(bytes,options),doc=new Y.Doc();try{for(const values of imported.rows){const row=new Y.Array<string>();row.push(values);doc.getArray('sheet').push([row]);}doc.getMap('sheetCellTypes').set('$default','text');const settings=doc.getMap('sheetImport');for(const [key,value]of Object.entries(imported.settings))settings.set(key,value);return {doc,update:Y.encodeStateAsUpdate(doc)};}catch(error){doc.destroy();throw error;}
}
export function validateSpreadsheetUpdate(current:Uint8Array|null,update:Uint8Array){
 const doc=new Y.Doc();try{
  if(current)Y.applyUpdate(doc,current);Y.applyUpdate(doc,update);validateOfficeComments(doc);if([...doc.share.keys()].some(key=>!['officeComments','officeCommentOperations','sheetColumns','sheet','sheetOperations','sheetImport','sheetCellTypes','sheetXlsx','sheetOrder','workbookSheets','sheetFormats','sheetMerges'].includes(key)))throw new Error('表格更新包含非法根结构');
  if(doc.share.has('sheetImport')&&!validDelimitedImportSettings(doc.getMap('sheetImport').toJSON()))throw new Error('表格导入设置无效');if(doc.share.has('sheetXlsx')&&!validXlsxSource(doc.getMap('sheetXlsx').toJSON()))throw new Error('XLSX 导入范围信息无效');
  if(doc.share.has('sheetOrder')){const ids=doc.getArray('sheetOrder').toArray(),entries=doc.getMap('workbookSheets');if(!ids.length||ids.length>MAX_WORKBOOK_SHEETS||new Set(ids).size!==ids.length||ids.some(id=>typeof id!=='string'||! /^[a-zA-Z0-9-]{1,128}$/.test(id))||[...entries.keys()].some(id=>!ids.includes(id)))throw new Error('工作表顺序无效');
   for(const id of ids){const value=entries.get(id as string);if(!(value instanceof Y.Map)||!validWorksheetName(value.get('name'))||!['text','auto'].includes(String(value.get('defaultType')))||[...value.keys()].some(key=>!['name','defaultType','sheet'].includes(key))||id!==FIRST_WORKSHEET&&!(value.get('sheet') instanceof Y.Array)||id===FIRST_WORKSHEET&&value.has('sheet'))throw new Error('工作表结构无效');}
   if(!ids.includes(FIRST_WORKSHEET)&&doc.getArray('sheet').length)throw new Error('工作簿含有未引用的旧工作表内容');
   const names=workbookSheets(doc).map(sheet=>sheet.name.toLowerCase());if(new Set(names).size!==names.length)throw new Error('工作表名称重复');
  }else if(doc.share.has('workbookSheets')&&doc.getMap('workbookSheets').size)throw new Error('工作簿缺少工作表顺序');
  let cells=0;for(const {sheet}of workbookSheets(doc)){if(!(sheet instanceof Y.Array)||sheet.length>10000)throw new Error('表格行数超过 1 万或结构无效');for(const value of sheet){if(!(value instanceof Y.Array))throw new Error('表格行结构无效');cells+=value.length;if(cells>MAX_SPREADSHEET_CELLS)throw new Error('表格单元格超过 20 万');if(value.length>16384)throw new Error('XLSX 列数超过 16384');for(const cell of value)if(typeof cell!=='string'||cell.length>10000)throw new Error('表格单元格内容无效');}}
  const sheetIds=new Set(workbookSheets(doc).map(sheet=>sheet.id));
  for(const [root,valid,limit]of [['sheetFormats',validSheetCellFormat,SHEET_FORMAT_LIMITS.formats],['sheetMerges',validSheetMerge,SHEET_FORMAT_LIMITS.merges]] as const)if(doc.share.has(root)){const map=doc.getMap(root);if(map._start||map.size>limit||[...map].some(([key,value])=>!sheetIds.has(parseSheetCellKey(key)?.sheet??'')||!valid(value)))throw new Error(root==='sheetFormats'?'表格单元格格式无效或超限':'表格合并区域无效或超限');}
  validateSheetColumns(doc);validateSpreadsheetCellTypes(doc);const receipts=doc.getMap('sheetOperations');if(receipts.size>10000||[...receipts.values()].some(value=>typeof value!=='string'||!/^[a-f0-9]{64}$/.test(value)))throw new Error('表格操作记录无效');
 }finally{doc.destroy();}
}
export {encodeXlsxWorkbook} from './office-xlsx';
