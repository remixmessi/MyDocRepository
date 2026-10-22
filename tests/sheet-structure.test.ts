import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {transformFormulaReferences,evaluateFormula} from '../packages/common/nbstore/src/spreadsheet-formula';
import {changeSheetStructure} from '../packages/frontend/apps/web/server/office-sheet-structure';
import {validateSpreadsheetUpdate} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {setSpreadsheetCellType,spreadsheetCellType} from '../packages/common/nbstore/src/spreadsheet-cell-type';
import {sheetColumnIds} from '../packages/common/nbstore/src/sheet-columns';
import {cellDraftKey,cellRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
const grid=(values:string[][])=>{const doc=new Y.Doc(),sheet=doc.getArray<Y.Array<string>>('sheet');for(const valuesRow of values){const row=new Y.Array<string>();row.push(valuesRow);sheet.push([row]);}return {doc,sheet};};
it('引用变换保留绝对引用和字符串，跨表引用调整，删除区域收缩，删除唯一引用形成 #REF!',()=>{
 const change={sheet:'数据',axis:'row' as const,index:1,delete:false};expect(transformFormulaReferences('=IF(A1="A2",$A$2,A2)','数据',change)).toBe('=IF((A1="A2"),$A$3,A3)');expect(transformFormulaReferences("='数据'!A2+B2",'汇总',change)).toBe("=('数据'!A3+B2)");expect(transformFormulaReferences('=SUM(A1:A3)','数据',{...change,delete:true})).toBe('=SUM(A1:A2)');expect(transformFormulaReferences('=A2','数据',{...change,delete:true})).toBe('=#REF!');expect(evaluateFormula('=IF(FALSE,#REF!,7)',[])).toBe('7');expect(evaluateFormula('=#REF!',[])).toBe('#REF!');expect(()=>transformFormulaReferences('=INDIRECT(A1)&A2','数据',change)).toThrow('安全引用变换');expect(transformFormulaReferences('=A2&TEXT(A2,"0")','数据',change)).toBe('=(A3&TEXT(A3,"0"))');
});
it('插删任意行列保留稳定单元格、文本类型和公式，支持备份校验',()=>{
 const {doc,sheet}=grid([['001','=A2+1'],['2','编号'],['3','=SUM(A1:A3)']]),row=sheet.get(1),column=sheetColumnIds(sheet)[0];setSpreadsheetCellType(doc,1,0,'text');const draft=cellDraftKey(sheet,1,0),base=cellRevision(sheet,1,0);
 changeSheetStructure(doc,'first','insertColumn',0,'add-column-12345678');expect(sheetColumnIds(sheet)[1]).toBe(column);expect(cellDraftKey(sheet,1,1)).toBe(draft);expect(cellRevision(sheet,1,1)).toBe(base);expect(spreadsheetCellType(doc,1,1)).toBe('text');expect(sheet.get(0).get(2)).toBe('=(B2+1)');
 changeSheetStructure(doc,'first','insertRow',1,'add-row-1234567890');expect(sheet.get(2)).toBe(row);expect(cellDraftKey(sheet,2,1)).toBe(draft);expect(spreadsheetCellType(doc,2,1)).toBe('text');expect(sheet.get(0).get(2)).toBe('=(B3+1)');changeSheetStructure(doc,'first','deleteRow',2,'delete-row-1234567');expect(sheet.get(0).get(2)).toBe('=(#REF!+1)');expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();changeSheetStructure(doc,'first','deleteColumn',0,'delete-column-12345');expect(sheetColumnIds(sheet)[0]).toBe(column);expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();doc.destroy();
});
it('未知公式阻止结构修改，文本类型的等号内容保持字节不变，旧数字列类型仍兼容',()=>{
 const {doc,sheet}=grid([['=INDIRECT("A"&A2)','001'],['2','3']]);const original=Y.encodeStateAsUpdate(doc);expect(()=>changeSheetStructure(doc,'first','insertRow',0,'block-op-123456789')).toThrow('安全引用变换');expect(Y.encodeStateAsUpdate(doc)).toEqual(original);setSpreadsheetCellType(doc,0,0,'text');const legacy=JSON.stringify([Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(sheet,0))),1]);doc.getMap('sheetCellTypes').set(legacy,'text');changeSheetStructure(doc,'first','insertColumn',0,'safe-op-1234567890');expect(sheet.get(0).get(1)).toBe('=INDIRECT("A"&A2)');expect(spreadsheetCellType(doc,0,2)).toBe('text');expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();doc.destroy();
});
