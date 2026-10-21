import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {strToU8,zipSync} from 'fflate';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {encodeXlsxWorkbook,validateSpreadsheetUpdate} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {exportXlsxFormula,exportXlsxText} from '../packages/frontend/apps/web/server/office-spreadsheet-export';
import {exportDelimitedText} from '../packages/frontend/apps/web/server/office-delimited-export';
import {validateOfficeBackupModel} from '../packages/frontend/apps/web/server/office-backup-model';
import {workbookSheets,workbookFormulaSnapshot} from '../packages/common/nbstore/src/office-workbook';
import {spreadsheetCellType} from '../packages/common/nbstore/src/spreadsheet-cell-type';
import {calculateFormulaWorkbookTyped} from '../packages/common/nbstore/src/spreadsheet-formula';
import {cellRevision,structureRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
import {submitSheetOperation} from '../packages/frontend/apps/web/server/office-sheet-operation';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';
function source(formula="'输入 数据'!A1*3"){const parts={
 'xl/workbook.xml':'<workbook xmlns:r="urn:rel"><sheets><sheet name="输入 数据" sheetId="9" r:id="first"/><sheet name="汇总" sheetId="2" r:id="second"/></sheets></workbook>',
 'xl/_rels/workbook.xml.rels':'<Relationships><Relationship Id="first" Type="office/worksheet" Target="worksheets/sheet9.xml"/><Relationship Id="second" Type="office/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
 'xl/worksheets/sheet9.xml':'<worksheet><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1" t="inlineStr"><is><t>001</t></is></c></row></sheetData></worksheet>',
 'xl/worksheets/sheet2.xml':'<worksheet><sheetData><row r="1"><c r="A1"><f>'+formula+'</f><v>9999</v></c><c r="B1"><f>\'输入 数据\'!B1</f><v>wrong</v></c></row></sheetData></worksheet>',
 };return zipSync(Object.fromEntries(Object.entries(parts).map(([name,value])=>[name,strToU8(value)])));}
it('工作簿按关系导入全部表及类型，跨表计算与三种副本往返、选表 CSV/TSV',async()=>{
 const parsed=encodeXlsxWorkbook(source()),doc=parsed.doc,pool=new OfficeParserPool();try{
  validateSpreadsheetUpdate(null,parsed.update);const sheets=workbookSheets(doc),{book}=workbookFormulaSnapshot(doc),types=(name:string,r:number,c:number)=>spreadsheetCellType(doc,r,c,sheets.find(sheet=>sheet.name===name)!.id);expect(sheets.map(sheet=>sheet.name)).toEqual(['输入 数据','汇总']);expect(spreadsheetCellType(doc,0,1,sheets[0].id)).toBe('text');expect(calculateFormulaWorkbookTyped(book,types)['汇总']).toEqual([[21,'001']]);
  for(const bytes of [exportXlsxFormula(parsed.update),exportXlsxFormula(parsed.update,true),exportXlsxText(parsed.update)]){const reopened=encodeXlsxWorkbook(bytes);try{expect(workbookSheets(reopened.doc).map(sheet=>sheet.name)).toEqual(['输入 数据','汇总']);expect(workbookSheets(reopened.doc)[1].sheet.get(0).toArray()).toEqual(["='输入 数据'!A1*3","='输入 数据'!B1"]);}finally{reopened.doc.destroy();}}
  for(const format of ['csv','tsv'] as const){const bytes=await pool.exportDelimited(parsed.update,format,'raw',undefined,'values',sheets[1].id);expect(new TextDecoder().decode(bytes)).toBe(format==='csv'?'21,001\r\n':'21\t001\r\n');}
  const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('spreadsheet',parsed.update));expect(workbookSheets(restored)).toHaveLength(2);expect(spreadsheetCellType(restored,0,0,'sheet-1')).toBe('formula');}finally{restored.destroy();}expect(Y.encodeStateAsUpdate(doc)).toEqual(parsed.update);
 }finally{doc.destroy();pool.close();}
});
it('跨表循环/错误拒绝，当前 CSV 只检查可达依赖；不使用陈旧缓存',()=>{
 const parsed=encodeXlsxWorkbook(source("'输入 数据'!A1/0"));try{expect(()=>exportXlsxFormula(parsed.update,true)).toThrow('汇总!A1 #DIV/0!');expect(new TextDecoder().decode(exportDelimitedText(parsed.update,'csv','raw','values','first'))).toBe('7,001\r\n');expect(()=>exportDelimitedText(parsed.update,'csv','raw','values','sheet-1')).toThrow('#DIV/0!');const first=parsed.doc.getArray<Y.Array<string>>('sheet').get(0);first.delete(0,1);first.insert(0,["='汇总'!A1"]);expect(()=>exportXlsxFormula(Y.encodeStateAsUpdate(parsed.doc),true)).toThrow('循环');expect(exportXlsxText(Y.encodeStateAsUpdate(parsed.doc)).length).toBeGreaterThan(100);}finally{parsed.doc.destroy();}
});
it('按工作表提交互不串写，类型与批量原子操作保持幂等',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-workbook-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),parsed=encodeXlsxWorkbook(source()),doc=parsed.doc;try{
  cloud.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:s','book',Buffer.from(parsed.update));const [first,second]=workbookSheets(doc),firstBase=cellRevision(first.sheet,0,0),secondBase=cellRevision(second.sheet,0,0),operation={operationId:crypto.randomUUID(),type:'setCell' as const,sheet:second.id,row:0,column:0,value:'二表独立',cellType:'text' as const,base:secondBase};Y.applyUpdate(doc,submitSheetOperation(cloud.database,'s','book','actor',operation).update);expect(first.sheet.get(0).get(0)).toBe('7');expect(second.sheet.get(0).get(0)).toBe('二表独立');expect(submitSheetOperation(cloud.database,'s','book','actor',operation).repeated).toBe(true);expect(()=>submitSheetOperation(cloud.database,'s','book','actor',{...operation,sheet:first.id})).toThrow('编号');
  Y.applyUpdate(doc,submitSheetOperation(cloud.database,'s','book','actor',{...operation,operationId:crypto.randomUUID(),sheet:first.id,base:firstBase,value:'9',cellType:'auto'}).update);expect(first.sheet.get(0).get(0)).toBe('9');expect(second.sheet.get(0).get(0)).toBe('二表独立');validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc));
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
it('工作表管理 CAS、删除改名引用保护、回执重试及配额验证',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-tab-ops-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),parsed=encodeXlsxWorkbook(source()),doc=parsed.doc;try{
  cloud.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:s','book',Buffer.from(parsed.update));const operation={operationId:crypto.randomUUID(),type:'addSheet' as const,sheet:'first',name:'空白',base:structureRevision(doc)};Y.applyUpdate(doc,submitSheetOperation(cloud.database,'s','book','actor',operation).update);expect(workbookSheets(doc)).toHaveLength(3);expect(submitSheetOperation(cloud.database,'s','book','actor',operation).repeated).toBe(true);expect(()=>submitSheetOperation(cloud.database,'s','book','actor',{...operation,operationId:crypto.randomUUID(),name:'过时'})).toThrow('已更新');
  const current=Buffer.from(cloud.database.load('cloud:s','book')!);for(const type of ['deleteSheet','renameSheet'] as const)expect(()=>submitSheetOperation(cloud.database,'s','book','actor',{operationId:crypto.randomUUID(),type,sheet:'first',name:type==='renameSheet'?'改名':undefined,base:structureRevision(doc)})).toThrow('引用');expect(cloud.database.load('cloud:s','book')).toEqual(current);
  const id='tab-'+operation.operationId,remove={operationId:crypto.randomUUID(),type:'deleteSheet' as const,sheet:id,base:structureRevision(doc)};Y.applyUpdate(doc,submitSheetOperation(cloud.database,'s','book','actor',remove).update);expect(workbookSheets(doc)).toHaveLength(2);expect(submitSheetOperation(cloud.database,'s','book','actor',remove).repeated).toBe(true);
  const order=doc.getArray('sheetOrder');order.push(['first']);expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('顺序');
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('工作簿预览共享预算；全局单元格配额不能按工作表绕过',async()=>{
 const {previewFormulaWorkbook,formulaMayReferenceSheet}=await import('../packages/common/nbstore/src/spreadsheet-formula');expect(previewFormulaWorkbook({'输入':[['3','001']],'汇总':[['=输入!A1*2','=输入!B1','=1/0']]},(name,_r,c)=>name==='输入'&&c===1?'text':'auto')['汇总']).toEqual([['6','001','#DIV/0!']]);
 const formula='=SUM('+Array(500).fill('1').join(',')+')',preview=previewFormulaWorkbook({a:[Array(250).fill(formula)],b:[Array(250).fill(formula)]});expect(preview.a[0][0]).toBe('500');expect(preview.b[0][249]).toBe('#LIMIT!');
 expect(formulaMayReferenceSheet("=SUM('输入'!A1:A100000)",'输入')).toBe(true);expect(formulaMayReferenceSheet('=UNKNOWN(', '输入')).toBe(true);expect(formulaMayReferenceSheet('="输入!A1"','输入')).toBe(false);
 const doc=new Y.Doc();try{doc.getMap('sheetCellTypes').set('$default','auto');for(const id of ['first','second']){const meta=new Y.Map<unknown>(),grid=id==='first'?doc.getArray<Y.Array<string>>('sheet'):new Y.Array<Y.Array<string>>();meta.set('name',id);meta.set('defaultType','auto');if(id!=='first')meta.set('sheet',grid);doc.getMap('workbookSheets').set(id,meta);doc.getArray('sheetOrder').push([id]);for(let i=0;i<7;i++){const row=new Y.Array<string>();row.push(Array(15000).fill(''));grid.push([row]);}}expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('20 万');}finally{doc.destroy();}
});
