import * as Y from 'yjs';
import {workbookSheet,workbookSheets} from '../../../../common/nbstore/src/office-workbook';
import {ensureSheetColumns,sheetColumnIds} from '../../../../common/nbstore/src/sheet-columns';
import {spreadsheetCellType,snapshotSheetCellTypes,clearSheetCellTypes,setSpreadsheetCellType} from '../../../../common/nbstore/src/spreadsheet-cell-type';
import {transformFormulaReferences,type FormulaStructureChange} from '../../../../common/nbstore/src/spreadsheet-formula';
export function changeSheetStructure(doc:Y.Doc,sheetId:string,type:string,index:number,operation:string){
 const selected=workbookSheet(doc,sheetId),sheet=selected.sheet,columnIds=sheetColumnIds(sheet),axis=type.endsWith('Column')?'column':'row',deleting=type.startsWith('delete'),length=axis==='row'?sheet.length:columnIds.length;
 if(!Number.isSafeInteger(index)||index<0||index>(deleting?length-1:length))throw new Error('行列位置无效或已删除');
 if(axis==='column'&&deleting&&length<=1)throw new Error('至少保留一列');
 if(!deleting&&length>=(axis==='row'?10000:16384))throw new Error('行列数量达到上限');
 const change:FormulaStructureChange={sheet:selected.name,axis,index,delete:deleting},rewrites:{row:Y.Array<string>;column:number;value:string}[]=[];let budget=0;
 // Plan the whole rewrite before mutating anything; unknown formula syntax aborts the transaction.
 for(const other of workbookSheets(doc))for(const [r,row]of other.sheet.toArray().entries())for(const [c,value]of row.toArray().entries()){
  if(other.id===sheetId&&deleting&&(axis==='row'?r:c)===index)continue;
  if(!value.trim().startsWith('=')||spreadsheetCellType(doc,r,c,other.id)==='text')continue;
  budget+=value.length;if(budget>200000)throw new Error('公式引用变换超过 20 万字符配额，请先简化公式');
  const next=transformFormulaReferences(value,other.name,change);if(next!==value)rewrites.push({row,column:c+(other.id===sheetId&&axis==='column'&&(deleting?c>index:c>=index)?deleting?-1:1:0),value:next});
 }
 const types=snapshotSheetCellTypes(doc,sheetId);clearSheetCellTypes(doc,sheetId);
 if(axis==='row'){
  if(deleting)sheet.delete(index,1);else{const row=new Y.Array<string>();row.push(Array(columnIds.length).fill(''));sheet.insert(index,[row]);}
 }else{
  const columns=ensureSheetColumns(doc,sheetId,sheet);
  if(deleting){columns.delete(index,1);for(const row of sheet)if(index<row.length)row.delete(index,1);}
  else{columns.insert(index,['column-'+operation]);for(const row of sheet){while(row.length<index)row.push(['']);row.insert(index,['']);}}
 }
 for(const entry of types){const n=entry[axis];if(deleting&&n===index)continue;const next=n+(deleting?n>index?-1:0:n>=index?1:0);setSpreadsheetCellType(doc,axis==='row'?next:entry.row,axis==='column'?next:entry.column,entry.type,sheetId);}
 for(const entry of rewrites){entry.row.delete(entry.column,1);entry.row.insert(entry.column,[entry.value]);}
}
