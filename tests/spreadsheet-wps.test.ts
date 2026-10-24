import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {submitSheetOperation} from '../packages/frontend/apps/web/server/office-sheet-operation';
import {validateSpreadsheetUpdate} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {exportXlsxText} from '../packages/frontend/apps/web/server/office-spreadsheet-export';
import {encodeXlsxWorkbook} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {readOfficePackage,readOfficeXml} from '../packages/frontend/apps/web/server/office-package';
import {sheetColumnIds,sheetRowIdentity} from '../packages/common/nbstore/src/sheet-columns';
import {cellRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
import {sheetCellFormat,sheetCellKey,sheetBorderCss,validSheetCellFormat} from '../packages/common/nbstore/src/spreadsheet-format';
import {sheetRangeKeys,sheetValidations,sheetValidationAt,validationAccepts,validationMessage,validSheetObject,type SheetValidation} from '../packages/common/nbstore/src/spreadsheet-objects';
import {zipSync,strToU8} from 'fflate';

const grid=(values:string[][])=>{const doc=new Y.Doc(),sheet=doc.getArray<Y.Array<string>>('sheet');for(const valuesRow of values){const row=new Y.Array<string>();row.push(valuesRow);sheet.push([row]);}return {doc,sheet};};

it('单元格格式扩展：删除线、字号、垂直对齐、自动换行与四边框线的校验和 CSS',()=>{
 expect(validSheetCellFormat({strike:true,fontSize:14.5,verticalAlign:'top',wrap:true,borderTop:'thin',borderBottom:'double:FF0000'})).toBe(true);
 for(const bad of [{fontSize:5},{fontSize:14.25},{verticalAlign:'center'},{wrap:false},{borderLeft:'wavy'},{borderLeft:'thin:red'},{borderLeft:'thin:FF0000:x'}])expect(validSheetCellFormat(bad)).toBe(false);
 expect(sheetBorderCss('thin')).toBe('1px solid #000000');expect(sheetBorderCss('double:1F4E79')).toBe('3px double #1F4E79');expect(sheetBorderCss('dashed')).toBe('1px dashed #000000');expect(sheetBorderCss(undefined)).toBeUndefined();
});

it('格式语义操作：每个目标可带自己的格式（外侧框线），与共享格式合并，非法目标格式被拒',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-sheet-wps-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),{doc,sheet}=grid([['a','b'],['c','d']]);try{
  cloud.database.append('cloud:space','book',Buffer.from(Y.encodeStateAsUpdate(doc)));const call=(input:Parameters<typeof submitSheetOperation>[4])=>{const result=submitSheetOperation(cloud.database,'space','book','actor',input);Y.applyUpdate(doc,result.update);return result;};
  const rows=sheet.toArray().map(sheetRowIdentity),cols=sheetColumnIds(sheet),target=(r:number,c:number,format?:object)=>({rowKey:rows[r],columnKey:cols[c],...(format?{format}:{})});
  // Outside borders of A1:B2: each corner gets two sides; the shared patch also sets a fill.
  call({operationId:crypto.randomUUID(),type:'setFormat',base:'',format:{fill:'FFF2CC'},targets:[target(0,0,{borderTop:'thin',borderLeft:'thin'}),target(0,1,{borderTop:'thin',borderRight:'thin'}),target(1,0,{borderBottom:'thin',borderLeft:'thin'}),target(1,1,{borderBottom:'medium:C00000',borderRight:'thin'})]});
  expect(sheetCellFormat(doc,'first',sheet,0,0)).toEqual({fill:'FFF2CC',borderTop:'thin',borderLeft:'thin'});expect(sheetCellFormat(doc,'first',sheet,1,1)).toEqual({fill:'FFF2CC',borderBottom:'medium:C00000',borderRight:'thin'});
  // An empty shared patch is allowed only when every target has its own format; null clears.
  call({operationId:crypto.randomUUID(),type:'setFormat',base:'',format:{},targets:[target(0,0,{borderTop:null,strike:true,fontSize:16,wrap:true,verticalAlign:'bottom'})]});expect(sheetCellFormat(doc,'first',sheet,0,0)).toEqual({fill:'FFF2CC',borderLeft:'thin',strike:true,fontSize:16,wrap:true,verticalAlign:'bottom'});
  for(const bad of [{format:{},targets:[target(0,0)]},{format:{fill:'FFF2CC'},targets:[target(0,0,{})]},{format:{fill:'FFF2CC'},targets:[{...target(0,0),extra:1}]},{format:{fill:'FFF2CC'},targets:[target(0,0,{borderTop:'wavy'})]},{format:{fontSize:200},targets:[target(0,0)]},{format:{unknown:true},targets:[target(0,0)]}])
   expect(()=>submitSheetOperation(cloud.database,'space','book','actor',{operationId:crypto.randomUUID(),type:'setFormat',base:'',...bad} as never)).toThrow(/格式无效|格式目标/);
 }finally{cloud.close();rmSync(dir,{recursive:true,force:true});doc.destroy();}
});

it('数据有效性：列表与数值/文本长度规则求值与提示，停止样式拒绝直接输入，警告样式允许，公式与批量粘贴不受限',()=>{
 const list={type:'list' as const,items:['北京','上海','广州']},whole={type:'whole' as const,operator:'between' as const,value:'1',value2:'100'};
 expect(validationAccepts(list,'上海')).toBe(true);expect(validationAccepts(list,'深圳')).toBe(false);expect(validationAccepts(list,'')).toBe(true);
 expect(validationAccepts(whole,'50')).toBe(true);expect(validationAccepts(whole,'0')).toBe(false);expect(validationAccepts(whole,'3.5')).toBe(false);expect(validationAccepts({type:'decimal',operator:'greaterThan',value:'0'},'0.5')).toBe(true);
 expect(validationAccepts({type:'textLength',operator:'lessOrEqual',value:'4'},'中文😀字')).toBe(true);expect(validationAccepts({type:'textLength',operator:'lessOrEqual',value:'4'},'中文😀字多')).toBe(false);
 const dir=mkdtempSync(join(tmpdir(),'mydoc-sheet-dv-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),{doc,sheet}=grid([['城市','数量'],['',''],['','']]);try{
  cloud.database.append('cloud:space','book',Buffer.from(Y.encodeStateAsUpdate(doc)));const call=(input:Parameters<typeof submitSheetOperation>[4])=>{const result=submitSheetOperation(cloud.database,'space','book','actor',input);Y.applyUpdate(doc,result.update);return result;};
  const cities:SheetValidation={kind:'validation',sheet:'first',range:sheetRangeKeys(sheet,{top:1,left:0,bottom:2,right:0}),rule:list,errorStyle:'stop'},counts:SheetValidation={kind:'validation',sheet:'first',range:sheetRangeKeys(sheet,{top:1,left:1,bottom:2,right:1}),rule:whole,errorStyle:'warning',message:'请填写 1–100'};
  call({operationId:crypto.randomUUID(),type:'putSheetObject',objectId:'dv-00000001',object:cities,base:''});call({operationId:crypto.randomUUID(),type:'putSheetObject',objectId:'dv-00000002',object:counts,base:''});
  expect(sheetValidations(doc,'first')).toHaveLength(2);expect(sheetValidationAt(doc,sheet,'first',2,0)?.id).toBe('dv-00000001');expect(sheetValidationAt(doc,sheet,'first',0,0)).toBeUndefined();
  expect(validationMessage(cities)).toBe('输入值必须是下拉列表中的一项：北京、上海、广州');expect(validationMessage(counts)).toBe('请填写 1–100');expect(validationMessage({...counts,message:undefined})).toBe('输入值必须是整数，且介于 1 与 100');
  const set=(row:number,column:number,value:string,cellType?:'text'|'auto'|'formula')=>call({operationId:crypto.randomUUID(),type:'setCell',row,column,value,base:cellRevision(sheet,row,column),...(cellType?{cellType}:{})});
  expect(()=>set(1,0,'深圳')).toThrow('下拉列表');set(1,0,'上海');set(1,1,'500');expect(sheet.get(1).toArray()).toEqual(['上海','500']);
  set(2,0,'=A2','formula');call({operationId:crypto.randomUUID(),type:'setCells',base:'',cells:[{row:2,column:0,value:'深圳',base:cellRevision(sheet,2,0),cellType:'auto'}]});expect(sheet.get(2).get(0)).toBe('深圳');
  for(const bad of [{...cities,rule:{type:'list',items:[]}},{...cities,rule:{type:'list',items:['a,b']}},{...cities,rule:{type:'list',items:['a','a']}},{...cities,rule:{type:'whole',operator:'between',value:'1'}},{...cities,rule:{type:'whole',operator:'equal',value:'1.5'}},{...cities,errorStyle:'info'},{...cities,message:''}])
   expect(validSheetObject(bad)).toBe(false);
  expect(()=>call({operationId:crypto.randomUUID(),type:'putSheetObject',objectId:'dv-00000001',object:{kind:'conditional',sheet:'first',range:cities.range,rule:{type:'blank',style:{bold:true}},priority:1},base:JSON.stringify(cities)})).toThrow('类型不能更改');
  const forged=new Y.Doc();Y.applyUpdate(forged,Y.encodeStateAsUpdate(doc));forged.getMap('sheetObjects').set('dv-bad-0001',{...cities,rule:{type:'list',items:[]}});expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(forged))).toThrow();forged.destroy();
 }finally{cloud.close();rmSync(dir,{recursive:true,force:true});doc.destroy();}
});

it('XLSX 往返：边框、删除线、字号、垂直对齐、换行与数据有效性；导入时不支持的有效性计数',()=>{
 const {doc,sheet}=grid([['城市','数量'],['北京','10'],['','']]),rows=sheet.toArray().map(sheetRowIdentity),cols=sheetColumnIds(sheet),formats=doc.getMap('sheetFormats');
 formats.set(sheetCellKey('first',rows[0],cols[0]),{bold:true,strike:true,fontSize:14,verticalAlign:'top',wrap:true,borderTop:'thin',borderBottom:'double:C00000',borderLeft:'dashed',borderRight:'medium'});
 const objects=doc.getMap('sheetObjects');objects.set('dv-00000001',{kind:'validation',sheet:'first',range:sheetRangeKeys(sheet,{top:1,left:0,bottom:2,right:0}),rule:{type:'list',items:['北京','上海 "浦东"']},errorStyle:'stop',message:'请选择城市',prompt:'从列表选择'});
 objects.set('dv-00000002',{kind:'validation',sheet:'first',range:sheetRangeKeys(sheet,{top:1,left:1,bottom:2,right:1}),rule:{type:'decimal',operator:'greaterOrEqual',value:'0'},errorStyle:'warning'});
 try{
  const bytes=exportXlsxText(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes),styles=readOfficeXml(files['xl/styles.xml']),sheetXml=readOfficeXml(files['xl/worksheets/sheet1.xml']);
  expect(styles).toContain('<border><left style="dashed"><color rgb="FF000000"/></left><right style="medium"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="double"><color rgb="FFC00000"/></bottom><diagonal/></border>');
  expect(styles).toContain('<font><b/><strike/><sz val="14"/><name val="Calibri"/></font>');expect(styles).toContain('<alignment vertical="top" wrapText="1"/>');
  expect(sheetXml).toContain('<dataValidations count="2"><dataValidation type="list" errorStyle="stop" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="输入值无效" error="请选择城市" promptTitle="提示" prompt="从列表选择" sqref="A2:A3"><formula1>&quot;北京,上海 &quot;&quot;浦东&quot;&quot;&quot;</formula1></dataValidation>');
  expect(sheetXml).toContain('operator="greaterThanOrEqual" errorStyle="warning"');expect(sheetXml.indexOf('</sheetData>')).toBeLessThan(sheetXml.indexOf('<dataValidations'));
  const back=encodeXlsxWorkbook(bytes);try{
   expect(sheetCellFormat(back.doc,'first',back.doc.getArray('sheet'),0,0)).toEqual({bold:true,strike:true,fontSize:14,verticalAlign:'top',wrap:true,borderTop:'thin',borderBottom:'double:C00000',borderLeft:'dashed',borderRight:'medium'});
   const imported=sheetValidations(back.doc,'first').map(([,value])=>value);expect(imported.map(value=>[value.rule,value.errorStyle,value.message??null])).toEqual([[{type:'list',items:['北京','上海 "浦东"']},'stop','请选择城市'],[{type:'decimal',operator:'greaterOrEqual',value:'0'},'warning',null]]);
  }finally{back.doc.destroy();}
  // Range-based lists and custom formulas are skipped and counted; Excel without showErrorMessage only warns.
  const minimal=(validations:string)=>zipSync({'xl/worksheets/sheet1.xml':strToU8(`<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData><dataValidations>${validations}</dataValidations></worksheet>`)});
  const parsed=encodeXlsxWorkbook(minimal('<dataValidation type="list" sqref="A1"><formula1>$C$1:$C$3</formula1></dataValidation><dataValidation type="custom" sqref="A1"><formula1>A1&gt;0</formula1></dataValidation><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>3</formula1></dataValidation>'));
  try{expect(parsed.doc.getMap('sheetXlsx').get('skippedValidations')).toBe(2);expect(sheetValidations(parsed.doc,'first').map(([,v])=>v.errorStyle)).toEqual(['warning']);}finally{parsed.doc.destroy();}
 }finally{doc.destroy();}
});

it('功能区命令：框线预设生成各单元格补丁，删除重复项上移保留行并清空末尾，分列检测覆盖与越界',async()=>{
 const {borderPatches,dedupePlan,splitPlan}=await import('../packages/frontend/core/src/spreadsheet-commands');
 const range={top:0,left:0,bottom:1,right:1};
 expect(borderPatches(range,'outside')).toEqual([{row:0,column:0,format:{borderTop:'thin',borderLeft:'thin'}},{row:0,column:1,format:{borderTop:'thin',borderRight:'thin'}},{row:1,column:0,format:{borderBottom:'thin',borderLeft:'thin'}},{row:1,column:1,format:{borderBottom:'thin',borderRight:'thin'}}]);
 expect(borderPatches(range,'bottom').map(item=>[item.row,item.column])).toEqual([[1,0],[1,1]]);expect(borderPatches(range,'none')[0].format.borderTop).toBeNull();expect(borderPatches(range,'thickOutside')[0].format.borderTop).toBe('thick');
 const grid=[['姓名','城市'],['张三','北京'],['李四','上海'],['张三 ','北京'],['王五','北京'],['张三','广州']];
 const all=dedupePlan(grid,{top:0,left:0,bottom:5,right:1},[0,1],true);expect(all.removed).toBe(1);expect(all.kept).toBe(4);
 const apply=(plan:{edits:{row:number;column:number;value:string}[]})=>{const copy=grid.map(row=>[...row]);for(const edit of plan.edits)copy[edit.row][edit.column]=edit.value;return copy;};
 expect(apply(all)).toEqual([['姓名','城市'],['张三','北京'],['李四','上海'],['王五','北京'],['张三','广州'],['','']]);
 expect(apply(dedupePlan(grid,{top:0,left:0,bottom:5,right:1},[1],true))).toEqual([['姓名','城市'],['张三','北京'],['李四','上海'],['张三','广州'],['',''],['','']]);
 const split=splitPlan([['a,b,c',''],['x,y','占用'],['单值','']],{top:0,left:0,bottom:2,right:0},0,',',3);expect(split.overwrite).toBe(1);expect(split.overflow).toBe(false);expect(split.columns).toBe(3);expect(split.edits).toContainEqual({row:0,column:2,value:'c'});
 expect(splitPlan([['a b  c']],{top:0,left:0,bottom:0,right:0},0,' ',2).overflow).toBe(true);
});
