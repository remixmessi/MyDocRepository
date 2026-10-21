import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {encodeDelimitedOriginal,validateSpreadsheetUpdate} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {submitSheetOperation} from '../packages/frontend/apps/web/server/office-sheet-operation';
import {cellRevision,structureRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
import {spreadsheetCellType,setSpreadsheetCellType,validateSpreadsheetCellTypes} from '../packages/common/nbstore/src/spreadsheet-cell-type';
import {evaluateFormula,detectFormulaCycles} from '../packages/common/nbstore/src/spreadsheet-formula';
import {validateOfficeBackupModel} from '../packages/frontend/apps/web/server/office-backup-model';
const options={delimiter:',' as const,quoteMode:'csv' as const};
it('导入默认文本，引用保留公式原文和前导零；旧表格继续自动计算',()=>{
 const parsed=encodeDelimitedOriginal(new TextEncoder().encode('=A1,001,=SUM(A1:B1)\n'),options),doc=parsed.doc,grid=doc.getArray<Y.Array<string>>('sheet').toArray().map(row=>row.toArray()),types=(_s:string,r:number,c:number)=>spreadsheetCellType(doc,r,c);
 try{expect(types('',0,0)).toBe('text');expect(evaluateFormula('=A1',grid,undefined,types)).toBe('=A1');expect(evaluateFormula('=B1',grid,undefined,types)).toBe('001');expect(evaluateFormula('=B1+1',grid,undefined,types)).toBe('2');expect(evaluateFormula('=SUM(A1:B1)',grid,undefined,types)).toBe('0');expect(detectFormulaCycles({当前工作表:grid},types)).toEqual([]);setSpreadsheetCellType(doc,0,0,'formula');expect(detectFormulaCycles({当前工作表:grid},types)).toContain('当前工作表!0:0');expect(evaluateFormula('=A1',grid,undefined,types)).toBe('#CYCLE!');}finally{doc.destroy();}
 expect(evaluateFormula('=A1',[['=1+1']])).toBe('2');
});
it('类型设置和引用严格校验，备份保留类型且不保留旧操作回执',()=>{
 const {doc}=encodeDelimitedOriginal(new TextEncoder().encode('=1+1,001\n'),options);try{setSpreadsheetCellType(doc,0,0,'formula');const bytes=Y.encodeStateAsUpdate(doc);validateSpreadsheetUpdate(null,bytes);const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('spreadsheet',bytes));expect(spreadsheetCellType(restored,0,0)).toBe('formula');expect(spreadsheetCellType(restored,0,1)).toBe('text');}finally{restored.destroy();}
 setSpreadsheetCellType(doc,0,1,'formula');expect(()=>validateSpreadsheetCellTypes(doc)).toThrow('=');setSpreadsheetCellType(doc,0,1,'text');doc.getMap('sheetCellTypes').set('foreign','text');expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('引用');doc.getMap('sheetCellTypes').delete('foreign');doc.getMap('sheetCellTypes').set('$default','formula');expect(()=>validateSpreadsheetCellTypes(doc)).toThrow('设置');}finally{doc.destroy();}
});
it('内容和类型同一 CAS 提交；仅类型变化冲突、回执幂等、非法类型回滚、删行清理',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-sheet-types-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),parsed=encodeDelimitedOriginal(new TextEncoder().encode('=1+1,001\n=A1,2\n'),options);try{
  cloud.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:space','doc:typed',Buffer.from(parsed.update));
  const operation={operationId:crypto.randomUUID(),type:'setCell' as const,row:0,column:0,value:'=1+1',cellType:'formula' as const,base:cellRevision(parsed.doc.getArray('sheet'),0,0)};
  const first=submitSheetOperation(cloud.database,'space','doc:typed','actor',operation);expect(first.repeated).toBe(false);const current=new Y.Doc();try{Y.applyUpdate(current,cloud.database.load('cloud:space','doc:typed')!);expect(spreadsheetCellType(current,0,0)).toBe('formula');expect(cellRevision(current.getArray('sheet'),0,0)).not.toBe(operation.base);expect(submitSheetOperation(cloud.database,'space','doc:typed','actor',operation).repeated).toBe(true);expect(()=>submitSheetOperation(cloud.database,'space','doc:typed','actor',{...operation,cellType:'text'})).toThrow('编号');expect(()=>submitSheetOperation(cloud.database,'space','doc:typed','actor',{...operation,operationId:crypto.randomUUID(),cellType:'text'})).toThrow('他人修改');
   const before=Buffer.from(cloud.database.load('cloud:space','doc:typed')!);expect(()=>submitSheetOperation(cloud.database,'space','doc:typed','actor',{...operation,operationId:crypto.randomUUID(),base:cellRevision(current.getArray('sheet'),0,0),cellType:'foreign'} as never)).toThrow('类型');expect(cloud.database.load('cloud:space','doc:typed')).toEqual(before);
   const last={operationId:crypto.randomUUID(),type:'setCell' as const,row:1,column:0,value:'=A1',cellType:'formula' as const,base:cellRevision(current.getArray('sheet'),1,0)};Y.applyUpdate(current,submitSheetOperation(cloud.database,'space','doc:typed','actor',last).update);
   Y.applyUpdate(current,submitSheetOperation(cloud.database,'space','doc:typed','actor',{operationId:crypto.randomUUID(),type:'deleteLastRow',base:structureRevision(current)}).update);validateSpreadsheetCellTypes(current);expect(current.getMap('sheetCellTypes').size).toBe(2);
   Y.applyUpdate(current,submitSheetOperation(cloud.database,'space','doc:typed','actor',{operationId:crypto.randomUUID(),type:'appendRow',base:structureRevision(current)}).update);expect(spreadsheetCellType(current,1,0)).toBe('text');
  }finally{current.destroy();}
 }finally{parsed.doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('服务端整表计算共用上下文，保留文本、布尔及错误样式字串，不修改快照',async()=>{
 const {calculateFormulaGrid}=await import('../packages/common/nbstore/src/spreadsheet-formula');const grid=[['=2','001','=A1*3','=B1','="#DIV/0!"','=1=1']],before=JSON.stringify(grid);
 expect(calculateFormulaGrid(grid,(_s,r,c)=>c===1?'text':'auto')).toEqual([['2','001','6','001','#DIV/0!','TRUE']]);expect(JSON.stringify(grid)).toBe(before);
 expect(calculateFormulaGrid([['=A1','#DIV/0!']],()=> 'text')).toEqual([['=A1','#DIV/0!']]);
});
it('整表计算拒绝循环、错误、不支持函数及全局计算配额，错误含单元格位置',async()=>{
 const {calculateFormulaGrid}=await import('../packages/common/nbstore/src/spreadsheet-formula');expect(()=>calculateFormulaGrid([['=1/0','=UNKNOWN(1)']])).toThrow('A1 #DIV/0!；B1 #NAME?');expect(()=>calculateFormulaGrid([['=A1']])).toThrow('循环依赖：A1');
 // Each SUM is individually cheap and has no references; only the shared evaluation budget is exceeded.
 const formula='=SUM('+Array(500).fill('1').join(',')+')';expect(()=>calculateFormulaGrid([Array(500).fill(formula)])).toThrow('#LIMIT!');
});

it('批量粘贴逐格 CAS 原子更新类型和内容，支持独立编辑、幂等重试与冲突整批回滚',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-batch-cells-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),parsed=encodeDelimitedOriginal(new TextEncoder().encode('a,b,c\nd,e,f\n'),options),doc=parsed.doc,sheet=doc.getArray<Y.Array<string>>('sheet');
 try{
  cloud.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:space','batch',Buffer.from(parsed.update));
  const cells=[{row:0,column:0,value:'001',cellType:'text' as const,base:cellRevision(sheet,0,0)},{row:1,column:1,value:'=1+1',cellType:'text' as const,base:cellRevision(sheet,1,1)}],operation={operationId:crypto.randomUUID(),type:'setCells' as const,base:'',cells};
  Y.applyUpdate(doc,submitSheetOperation(cloud.database,'space','batch','actor',{operationId:crypto.randomUUID(),type:'setCell',row:0,column:2,value:'独立编辑',base:cellRevision(sheet,0,2)}).update);
  Y.applyUpdate(doc,submitSheetOperation(cloud.database,'space','batch','actor',operation).update);expect(sheet.toArray().map(r=>r.toArray())).toEqual([['001','b','独立编辑'],['d','=1+1','f']]);expect(spreadsheetCellType(doc,1,1)).toBe('text');expect(submitSheetOperation(cloud.database,'space','batch','actor',operation).repeated).toBe(true);
  expect(()=>submitSheetOperation(cloud.database,'space','batch','actor',{...operation,cells:[{...cells[0],value:'changed'},cells[1]]})).toThrow('编号');
  const fresh={...operation,operationId:crypto.randomUUID(),cells:cells.map(cell=>({...cell,value:'不应写入',base:cellRevision(sheet,cell.row,cell.column)}))};
  Y.applyUpdate(doc,submitSheetOperation(cloud.database,'space','batch','actor',{operationId:crypto.randomUUID(),type:'setCell',row:1,column:1,value:'=1+1',cellType:'formula',base:cellRevision(sheet,1,1)}).update);
  const before=Buffer.from(cloud.database.load('cloud:space','batch')!);expect(()=>submitSheetOperation(cloud.database,'space','batch','actor',fresh)).toThrow('第 2 行第 2 列已更新');expect(cloud.database.load('cloud:space','batch')).toEqual(before);
  const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('spreadsheet',before));expect(restored.getArray<Y.Array<string>>('sheet').get(0).get(0)).toBe('001');expect(spreadsheetCellType(restored,1,1)).toBe('formula');}finally{restored.destroy();}
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
it('批量粘贴拒绝重复、越界、非法类型、配额并在数据库失败时回滚',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-batch-reject-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),parsed=encodeDelimitedOriginal(new TextEncoder().encode('a,b\n'),options),doc=parsed.doc,sheet=doc.getArray<Y.Array<string>>('sheet');
 try{
  cloud.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:space','batch',Buffer.from(parsed.update));const before=Buffer.from(parsed.update),cell={row:0,column:0,value:'new',base:cellRevision(sheet,0,0)},operation={operationId:crypto.randomUUID(),type:'setCells' as const,base:'',cells:[cell]};
  for(const cells of [[],Array(501).fill(cell),[cell,cell],[cell,{...cell,column:2}],[{...cell,cellType:'formula'}],[{...cell,value:'x'.repeat(10001)}],[{...cell,cellType:'invalid'}],[{...cell,row:-1}],[{...cell,base:7}]]){expect(()=>submitSheetOperation(cloud.database,'space','batch','actor',{...operation,cells} as never)).toThrow();expect(cloud.database.load('cloud:space','batch')).toEqual(before);}
  cloud.database.db.exec("CREATE TRIGGER fail_batch BEFORE UPDATE ON documents BEGIN SELECT RAISE(ABORT,'batch write failed'); END");expect(()=>submitSheetOperation(cloud.database,'space','batch','actor',operation)).toThrow('batch write failed');expect(cloud.database.load('cloud:space','batch')).toEqual(before);cloud.database.db.exec('DROP TRIGGER fail_batch');expect(submitSheetOperation(cloud.database,'space','batch','actor',operation).repeated).toBe(false);
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
