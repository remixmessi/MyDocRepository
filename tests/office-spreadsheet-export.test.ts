import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {exportXlsxText} from '../packages/frontend/apps/web/server/office-spreadsheet-export';
import {encodeXlsxWorkbook} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {readOfficePackage,readOfficeXml} from '../packages/frontend/apps/web/server/office-package';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';
function snapshot(values:string[][]){const doc=new Y.Doc();for(const valuesRow of values){const row=new Y.Array<string>();row.push(valuesRow);doc.getArray('sheet').push([row]);}return doc;}
it('XLSX 文本副本保留精确字串、空格、空格子、Unicode 和公式原文',()=>{
 const values=[['00123','12345678901234567890','=SUM(A1:A2)','+1','-2','@test'],[' 中文😀 &<>"\'\t\n\r ','&amp;quot;','_x000A_','_x005F_x0041_','','末尾']];const doc=snapshot(values);
 try{const bytes=exportXlsxText(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes),xml=readOfficeXml(files['xl/worksheets/sheet1.xml']);expect(Object.keys(files)).toHaveLength(7);expect(files['xl/styles.xml']).toBeDefined();expect(xml).not.toMatch(/<f[\s>]/);expect(xml).toContain('t="s"');expect(readOfficeXml(files['xl/workbook.xml'])).toContain('name="当前工作表"');const parsed=encodeXlsxWorkbook(bytes);try{expect(parsed.doc.getArray<Y.Array<string>>('sheet').toArray().map(row=>row.toArray())).toEqual(values);}finally{parsed.doc.destroy();}}finally{doc.destroy();}
});
it('列地址跨 Z/AA，空表可读，非法字符、列上限、超大 XML 明确拒绝',()=>{
 const doc=snapshot([Array.from({length:27},(_,i)=>String(i))]);try{expect(readOfficeXml(readOfficePackage(exportXlsxText(Y.encodeStateAsUpdate(doc)))['xl/worksheets/sheet1.xml'])).toContain('r="AA1"');}finally{doc.destroy();}
 const empty=snapshot([]);try{expect(exportXlsxText(Y.encodeStateAsUpdate(empty)).length).toBeGreaterThan(100);}finally{empty.destroy();}
 for(const value of ['\u0000','\uffff']){const bad=snapshot([[value]]);try{expect(()=>exportXlsxText(Y.encodeStateAsUpdate(bad))).toThrow('字符');}finally{bad.destroy();}}
 const wide=snapshot([Array(16385).fill('')]);try{expect(()=>exportXlsxText(Y.encodeStateAsUpdate(wide))).toThrow('列数');}finally{wide.destroy();}
 const large=snapshot([Array.from({length:220},(_,i)=>String(i)+'<'.repeat(9997))]);try{expect(()=>exportXlsxText(Y.encodeStateAsUpdate(large))).toThrow('8 MB');}finally{large.destroy();}
});
it('隔离子进程生成 XLSX 并可重新解析，不改变输入协作状态',async()=>{
 const doc=snapshot([['001','=1+1','_x000D_']]),pool=new OfficeParserPool();try{const before=Y.encodeStateAsUpdate(doc),bytes=await pool.exportXlsxText(before),update=await pool.parse('xlsx',bytes),parsed=new Y.Doc();try{Y.applyUpdate(parsed,update);expect(parsed.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(['001','=1+1','_x000D_']);expect(Y.encodeStateAsUpdate(doc)).toEqual(before);}finally{parsed.destroy();}}finally{doc.destroy();pool.close();}
});

it('CSV/TSV 明确转义公式样式文本，保留 Unicode、空列、引号及字段内换行',async()=>{
 const {encodeDelimitedText,delimitedText}=await import('../packages/common/nbstore/src/office-delimited-export');const {parseDelimited}=await import('../packages/common/nbstore/src/office');
 const rows=[['001','12345678901234567890','=1+1','  =1','\t正文','+2','-3','@x'],['中文😀','a,b','a\tb','a"b','a\r\nb','','','末尾']];
 for(const format of ['csv','tsv'] as const)for(const mode of ['raw','escaped'] as const){const bytes=encodeDelimitedText(rows,format,mode);expect([...bytes.slice(0,3)]).toEqual([239,187,191]);const text=new TextDecoder().decode(bytes);expect(text.endsWith('\r\n')).toBe(true);expect(parseDelimited(text,format==='csv'?',':'\t')).toEqual(rows.map(row=>row.map(value=>delimitedText(value,mode))));}
 for(const value of ['=1','+1','-1','@x','\t文本','\r文本','\n文本','  =1','\u0000=1','＝1'])expect(delimitedText(value,'escaped')).toBe("'"+value);
 for(const value of ["'=1",'001','12345678901234567890','中文😀','文字=1',''])expect(delimitedText(value,'escaped')).toBe(value);
 expect(new TextDecoder().decode(encodeDelimitedText([['a'],['b','']], 'csv','raw'))).toBe('a,\r\nb,\r\n');expect(new TextDecoder().decode(encodeDelimitedText([], 'csv','escaped'))).toBe('');
 expect(()=>encodeDelimitedText(Array.from({length:10001},()=>['']), 'csv','raw')).toThrow('1 万行');expect(()=>encodeDelimitedText([Array(100001).fill(''),Array(100001).fill('')], 'csv','raw')).toThrow('20 万');expect(()=>encodeDelimitedText([['a'.repeat(16*1024*1024)]], 'csv','raw')).toThrow('16 MB');
});
it('CSV/TSV 导出在受限子进程运行，原模型与回执保持不变',async()=>{
 const {parseDelimited}=await import('../packages/common/nbstore/src/office');const doc=snapshot([['=1+1','001','中文\t换行\n末尾']]),pool=new OfficeParserPool();try{const before=Y.encodeStateAsUpdate(doc);for(const format of ['csv','tsv'] as const){const bytes=await pool.exportDelimited(before,format,'escaped');expect(parseDelimited(new TextDecoder().decode(bytes),format==='csv'?',':'\t')[0]).toEqual(["'=1+1",'001','中文\t换行\n末尾']);}expect(Y.encodeStateAsUpdate(doc)).toEqual(before);}finally{pool.close();doc.destroy();}
});

it('CSV/TSV 计算结果从隔离快照生成，错误可显式改导出原文',async()=>{
 const {parseDelimited}=await import('../packages/common/nbstore/src/office'),{setSpreadsheetCellType}=await import('../packages/common/nbstore/src/spreadsheet-cell-type');const doc=snapshot([['=1+1','001','=B1','="#DIV/0!"']]),pool=new OfficeParserPool();try{
  setSpreadsheetCellType(doc,0,1,'text');const bytes=Y.encodeStateAsUpdate(doc);
  for(const format of ['csv','tsv'] as const){const result=await pool.exportDelimited(bytes,format,'raw',undefined,'values');expect(parseDelimited(new TextDecoder().decode(result),format==='csv'?',':'\t')[0]).toEqual(['2','001','001','#DIV/0!']);}expect(Y.encodeStateAsUpdate(doc)).toEqual(bytes);
  const row=doc.getArray<Y.Array<string>>('sheet').get(0);row.delete(0,1);row.insert(0,['=1/0']);const broken=Y.encodeStateAsUpdate(doc);await expect(pool.exportDelimited(broken,'csv','raw',undefined,'values')).rejects.toMatchObject({status:400,message:expect.stringContaining('A1 #DIV/0!')});expect(parseDelimited(new TextDecoder().decode(await pool.exportDelimited(broken,'csv','raw')))[0][0]).toBe('=1/0');
 }finally{doc.destroy();pool.close();}
});

it('XLSX 公式副本写入公式和数值，文本不执行；不伪造缓存并要求重算',async()=>{
 const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export'),{setSpreadsheetCellType}=await import('../packages/common/nbstore/src/spreadsheet-cell-type');const doc=snapshot([['2','=A1*3','001','=1+1','="_x000A_"']]);try{setSpreadsheetCellType(doc,0,2,'text');setSpreadsheetCellType(doc,0,3,'text');const before=Y.encodeStateAsUpdate(doc),files=readOfficePackage(exportXlsxFormula(before)),xml=readOfficeXml(files['xl/worksheets/sheet1.xml']);expect(xml).toContain('<c r="A1"><v>2</v></c>');expect(xml).toContain('<c r="B1"><f>A1*3</f></c>');expect(xml).toContain('<f>&quot;_x000A_&quot;</f>');expect(xml).not.toContain('<f>A1*3</f><v>');expect(readOfficeXml(files['xl/workbook.xml'])).toContain('fullCalcOnLoad="1"');const reopened=encodeXlsxWorkbook(exportXlsxFormula(before));try{expect(reopened.doc.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(['2','=A1*3','001','=1+1','="_x000A_"']);const {spreadsheetCellType}=await import('../packages/common/nbstore/src/spreadsheet-cell-type');expect(spreadsheetCellType(reopened.doc,0,1)).toBe('formula');expect(spreadsheetCellType(reopened.doc,0,3)).toBe('text');}finally{reopened.doc.destroy();}expect(Y.encodeStateAsUpdate(doc)).toEqual(before);}finally{doc.destroy();}
});
it('公式副本拒绝错误、循环、超精度自动数字，仍可选择文本副本',async()=>{
 const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export');for(const [value,message]of [['=1/0','#DIV/0!'],['=A1','循环'],['=UNKNOWN(1)','#NAME?'],['12345678901234567890','精度'],['1e-999','精度']]){const doc=snapshot([[value]]);try{expect(()=>exportXlsxFormula(Y.encodeStateAsUpdate(doc))).toThrow(message);expect(exportXlsxText(Y.encodeStateAsUpdate(doc)).length).toBeGreaterThan(100);}finally{doc.destroy();}}
 const pool=new OfficeParserPool(),doc=snapshot([['=2+3']]);try{const bytes=await pool.exportXlsxFormula(Y.encodeStateAsUpdate(doc)),parsed=encodeXlsxWorkbook(bytes);try{expect(parsed.doc.getArray<Y.Array<string>>('sheet').get(0).get(0)).toBe('=2+3');}finally{parsed.doc.destroy();}}finally{pool.close();doc.destroy();}
});

it('可选公式缓存保留数值、布尔、空字符串和文本类型，隔离导出不修改快照',async()=>{
 const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export');
 const doc=snapshot([['=2+3','=1<2','="001"','="TRUE"','="#DIV/0!"','=""','="_x000A_"','=A2','="中文 &<>"']]),pool=new OfficeParserPool();
 try{
  const before=Y.encodeStateAsUpdate(doc),bytes=exportXlsxFormula(before,true),xml=readOfficeXml(readOfficePackage(bytes)['xl/worksheets/sheet1.xml']);
  expect(xml).toContain('<c r="A1"><f>2+3</f><v>5</v></c>');expect(xml).toContain('<c r="B1" t="b"><f>1&lt;2</f><v>1</v></c>');
  for(const [ref,value]of [['C1','001'],['D1','TRUE'],['E1','#DIV/0!'],['F1',''],['G1','_x000A_'],['I1','中文 &amp;&lt;&gt;']])expect(xml).toMatch(new RegExp('<c r="'+ref+'" t="str"><f>.*?</f><v>'+value+'</v></c>'));
  expect(xml).toContain('<c r="H1"><f>A2</f><v>0</v></c>');
  expect(readOfficeXml(readOfficePackage(await pool.exportXlsxFormula(before,undefined,true))['xl/worksheets/sheet1.xml'])).toBe(xml);
  const parsed=encodeXlsxWorkbook(bytes);try{expect(parsed.doc.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(doc.getArray<Y.Array<string>>('sheet').get(0).toArray());}finally{parsed.doc.destroy();}
  expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
  doc.getArray<Y.Array<string>>('sheet').get(0).delete(0,1);doc.getArray<Y.Array<string>>('sheet').get(0).insert(0,['=1/0']);
  await expect(pool.exportXlsxFormula(Y.encodeStateAsUpdate(doc),undefined,true)).rejects.toMatchObject({status:400,message:expect.stringContaining('A1 #DIV/0!')});
 }finally{pool.close();doc.destroy();}
});

it('XLSX 和 CSV/TSV 以导入表名计算本表引用，副本保留名称且不改公式',async()=>{
 const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export'),{exportDelimitedText}=await import('../packages/frontend/apps/web/server/office-delimited-export');
 const {calculateFormulaGrid,evaluateFormula}=await import('../packages/common/nbstore/src/spreadsheet-formula');
 for(const name of ["销售 O'Brien & 明细",'__proto__','当前工作表','Sheet1']){
  const quoted="'"+name.replaceAll("'","''")+"'",values=[['2','='+quoted+'!A1*3','=SUM('+quoted+'!A1:B1)']],doc=snapshot(values);
  try{
   for(const [key,value]of Object.entries({name,sheetCount:2,hasStyles:false,textConversions:false}))doc.getMap('sheetXlsx').set(key,value);
   const before=Y.encodeStateAsUpdate(doc);expect(calculateFormulaGrid(values,undefined,name)[0]).toEqual(['2','6','8']);expect(evaluateFormula(values[0][1],values,{[name]:values})).toBe('6');
   for(const format of ['csv','tsv'] as const){const {parseDelimited}=await import('../packages/common/nbstore/src/office');expect(parseDelimited(new TextDecoder().decode(exportDelimitedText(before,format,'raw','values')),format==='csv'?',':'\t')[0]).toEqual(['2','6','8']);}
   for(const bytes of [exportXlsxText(before),exportXlsxFormula(before),exportXlsxFormula(before,true)]){const reopened=encodeXlsxWorkbook(bytes);try{expect(reopened.doc.getMap('sheetXlsx').get('name')).toBe(name);expect(reopened.doc.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(values[0]);}finally{reopened.doc.destroy();}}
   expect(readOfficeXml(readOfficePackage(exportXlsxFormula(before,true))['xl/worksheets/sheet1.xml'])).toContain('<v>6</v>');expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
  }finally{doc.destroy();}
 }
});
it('命名本表循环仍被拒绝，不把其他表或旧默认名称当作本表',async()=>{
 const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export');
 for(const [formula,error]of [["='销售'!A1",'循环'],["='其他表'!A1",'#REF!'],["='当前工作表'!A1",'#REF!']]){const doc=snapshot([[formula]]);try{for(const [key,value]of Object.entries({name:'销售',sheetCount:2,hasStyles:false,textConversions:false}))doc.getMap('sheetXlsx').set(key,value);expect(()=>exportXlsxFormula(Y.encodeStateAsUpdate(doc),true)).toThrow(error);expect(exportXlsxText(Y.encodeStateAsUpdate(doc)).length).toBeGreaterThan(100);}finally{doc.destroy();}}
});
it('旧模型默认表名保持兼容，不合法导出名称明确拒绝而非静默改名',async()=>{
 const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export');
 const legacy=snapshot([['2',"='当前工作表'!A1"]]);try{const files=readOfficePackage(exportXlsxFormula(Y.encodeStateAsUpdate(legacy),true));expect(readOfficeXml(files['xl/workbook.xml'])).toContain('name="当前工作表"');expect(readOfficeXml(files['xl/worksheets/sheet1.xml'])).toContain('<v>2</v>');expect(legacy.share.has('sheetXlsx')).toBe(false);}finally{legacy.destroy();}
 for(const name of ['a'.repeat(32),'销售/明细',"'销售","销售'",'销售[1]']){const doc=snapshot([['2']]);try{for(const [key,value]of Object.entries({name,sheetCount:1,hasStyles:false,textConversions:false}))doc.getMap('sheetXlsx').set(key,value);expect(()=>exportXlsxText(Y.encodeStateAsUpdate(doc))).toThrow('工作表名称');expect(()=>exportXlsxFormula(Y.encodeStateAsUpdate(doc),true)).toThrow('工作表名称');}finally{doc.destroy();}}
});

it('单元格讨论导出为 Excel 批注：按单元格合并线程与回复、标注已解决，无定位的讨论不写入，正文不变',async()=>{
 const {workbookSheets}=await import('../packages/common/nbstore/src/office-workbook');
 const {sheetColumnIds,sheetRowIdentity}=await import('../packages/common/nbstore/src/sheet-columns');
 const doc=snapshot([['名称','数量'],['苹果','3'],['梨','5']]),sheet=workbookSheets(doc)[0],rows=sheet.sheet.toArray(),columns=sheetColumnIds(sheet.sheet),comments=doc.getMap('officeComments');
 const anchor=(r:number,c:number,label:string)=>({kind:'cell' as const,sheet:sheet.id,row:sheetRowIdentity(rows[r]),column:columns[c],label});
 const message=(id:string,name:string,body:string,created:number)=>({id,author:'u'+id,authorName:name,body,created});
 comments.set('a',{id:'a',resolved:false,updated:2,anchor:anchor(1,1,'B2'),messages:[message('m1','王小明','数量 <待核> & 确认',1),message('m2','Lisa Chen','已确认',2)]});
 comments.set('b',{id:'b',resolved:true,updated:4,anchor:anchor(1,1,'B2'),messages:[message('m3','王小明','单位是箱',3)]});
 comments.set('c',{id:'c',resolved:false,updated:5,anchor:anchor(2,0,'A3'),messages:[message('m4','Lisa Chen','名称改为雪梨',5)]});
 comments.set('d',{id:'d',resolved:false,updated:6,messages:[message('m5','王小明','整表意见',6)]});
 try{
  const bytes=exportXlsxText(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes),notes=readOfficeXml(files['xl/comments1.xml']);
  expect(readOfficeXml(files['xl/worksheets/sheet1.xml'])).toContain('<legacyDrawing r:id="rIdVml"/>');expect(readOfficeXml(files['[Content_Types].xml'])).toContain('/xl/comments1.xml');expect(files['xl/drawings/vmlDrawing1.vml']).toBeDefined();
  expect([...notes.matchAll(/<comment ref="([A-Z]+\d+)"/g)].map(m=>m[1])).toEqual(['B2','A3']);expect(notes).toContain('<author>王小明</author>');expect(notes).toContain('<author>Lisa Chen</author>');
  const text=notes.replace(/<[^>]+>/g,'');expect(text).toContain('王小明：数量 &lt;待核&gt; &amp; 确认');expect(text).toContain('Lisa Chen：已确认');expect(text).toContain('[已解决] 王小明：单位是箱');expect(text).not.toContain('整表意见');
  const imported=encodeXlsxWorkbook(bytes);try{expect(imported.doc.getArray<Y.Array<string>>('sheet').toArray().map(row=>row.toArray())).toEqual([['名称','数量'],['苹果','3'],['梨','5']]);}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
