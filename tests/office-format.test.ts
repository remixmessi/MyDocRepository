import {describe,expect,it} from 'vitest';
import {decodeOfficeText,encodeDelimited,encodeOfficeText,exportOfficeText,inspectOfficeText,officeKindFromFilename,parseDelimited,prepareOfficeTextImport,validTextSettings} from '../packages/common/nbstore/src/office';
import {detectFormulaCycles,evaluateFormula} from '../packages/common/nbstore/src/spreadsheet-formula';

describe('办公文件格式基础工具',()=>{
 it('识别新增办公资源类型',()=>{
  expect(officeKindFromFilename('演示.PPTX')?.kind).toBe('presentation');
  expect(officeKindFromFilename('数据.tsv')?.kind).toBe('spreadsheet');
  expect(officeKindFromFilename('说明.TXT')?.kind).toBe('text');
  expect(officeKindFromFilename('未知.bin')).toBeUndefined();
 });
 it('保留 UTF-16 编码和 BOM',()=>{
  const bytes=encodeOfficeText('中文\nline','utf-16le',true);
  expect(decodeOfficeText(bytes)).toEqual({text:'中文\nline',encoding:'utf-16le',bom:true});
  expect(decodeOfficeText(encodeOfficeText('纯文本','utf-8',true))).toEqual({text:'纯文本',encoding:'utf-8',bom:true});
 });
 it('解析引号内逗号、换行和双引号',()=>{
  const source=encodeDelimited([['name','note'],['a,b','line1\nline2'],['say "hi"','']]);
  expect(parseDelimited(source)).toEqual([['name','note'],['a,b','line1\nline2'],['say "hi"','']]);
  expect(parseDelimited('a\tb\n1\t2\n','\t')).toEqual([['a','b'],['1','2']]);
 });
 it('拒绝未闭合的字段引用',()=>{expect(()=>parseDelimited('"broken,x')).toThrow('引号未闭合');});
	it('识别混合换行并要求明确规范化策略',()=>{
		const bytes=new TextEncoder().encode('a\r\nb\nc\rd');
		expect(inspectOfficeText(bytes).mixed).toBe(true);
		expect(()=>prepareOfficeTextImport(bytes)).toThrow('混合换行');
		expect(prepareOfficeTextImport(bytes,'utf-8','crlf').settings.lineEnding).toBe('crlf');
	});
	it('按设置导出换行、编码和 BOM',()=>{
		const bytes=exportOfficeText('a\nb',{encoding:'utf-16be',bom:true,lineEnding:'cr'});
		expect(decodeOfficeText(bytes)).toEqual({text:'a\rb',encoding:'utf-16be',bom:true});
	});
	it('拒绝错误 BOM、非法代理字符和超限文件',()=>{
		expect(()=>decodeOfficeText(encodeOfficeText('x','utf-16le',true),'utf-16be')).toThrow('BOM');
		expect(()=>encodeOfficeText('\ud800')).toThrow('代理字符');
		expect(()=>inspectOfficeText(new Uint8Array(20*1024*1024+1))).toThrow('20 MB');
		expect(validTextSettings({encoding:'utf-8',bom:false,lineEnding:'lf'})).toBe(true);
		expect(validTextSettings({encoding:'utf-8',bom:false})).toBe(false);
	});
	it('计算首批表格公式而不执行脚本',()=>{
		const grid=[['2','3'],['4','=SUM(A1:B1)'],['','=IF(A1>1,"是","否")']];
		expect(evaluateFormula('=SUM(A1:B2)',grid)).toBe('14');
		expect(evaluateFormula('=AVERAGE(A1:B1)',grid)).toBe('2.5');
		expect(evaluateFormula('=IF(A1>1,"是","否")',grid)).toBe('是');
		expect(evaluateFormula('=ROUND(1.235,2)',grid)).toBe('1.24');
		expect(evaluateFormula('=BAD(A1)',grid)).toBe('#NAME?');
		expect(evaluateFormula("='汇总'!A1+1",grid,{'汇总':[['4']]})).toBe('5');
		expect(evaluateFormula("=SUM('汇总'!A1:A2)",grid,{'汇总':[['4'],['6']]})).toBe('10');
		expect(detectFormulaCycles({A:[['=B!A1']],B:[['=A!A1']]})).toEqual(['A!0:0']);
		expect(detectFormulaCycles({A:[['=A1']]})).toEqual(['A!0:0']);
	});
});


describe('公式解析与依赖回归',()=>{
 it.each([
  ['=10-3-2','5'],['=12/3/2','2'],['=2*3+4','10'],['=2+3*4','14'],
  ['=(2+3)*4','20'],['=-2*-3','6'],['=-(2+3)','-5'],['=1e2+0.5','100.5'],
  ['=SUM(1,MAX(2,3))+4','8'],['=IF(2>1,4,1/0)+2','6'],
  ['=IF(FALSE,1/0,7)','7'],['=IF(TRUE,7,BAD(1))','7'],
  ['=IF(TRUE,"a,b","x")','a,b'],['=IF(TRUE,"a""b","x")','a"b'],
  ['=IF(TRUE,"+","-")','+'],['=ROUND(-1.235,2)','-1.24'],
  ['=ROUND(125,-1)','130'],['=1/0','#DIV/0!'],['=MISSING(1)','#NAME?'],
  ['=1+','#VALUE!'],['=(1+2','#VALUE!'],['=SUM(1))','#VALUE!'],
  ['=IF(TRUE,1,2,3)','#VALUE!'],['=1e309','#NUM!'],['=A0','#NAME?'],
 ])('计算 %s 得到 %s',(formula,result)=>expect(evaluateFormula(formula,[])).toBe(result));
 it('递归计算引用，保留字符串，不把错误引用转换为零',()=>{
  const grid=[['2','=A1+3','=B1*2','hello','']];
  expect(evaluateFormula('=C1',grid)).toBe('10');expect(evaluateFormula('=$A$1+1',grid)).toBe('3');
  expect(evaluateFormula('=D1',grid)).toBe('hello');expect(evaluateFormula('=D1+1',grid)).toBe('#VALUE!');
  expect(evaluateFormula('=Missing!A1',grid)).toBe('#REF!');
  expect(evaluateFormula('=XFE1',grid)).toBe('#REF!');
  expect(evaluateFormula('=COUNT(A1:E1)',grid)).toBe('3');expect(evaluateFormula('=AVERAGE(D1:E1)',grid)).toBe('#DIV/0!');
 });
 it('跨表公式递归解析，支持含空格、符号和转义单引号的工作表名',()=>{
  const base=[['3']],other=[['=Base!A1*2']];const book={'Base':base,"O'Brien-季度":other};
  expect(evaluateFormula("='O''Brien-季度'!A1+2",base,book)).toBe('8');
  expect(evaluateFormula('=base!A1',base,book)).toBe('3');
 });
 it('检测同表、跨表和范围循环；普通文本与字符串中的地址不产生依赖',()=>{
  expect(evaluateFormula('=A1',[['=A1']])).toBe('#CYCLE!');
  expect(evaluateFormula('=A1',[['=B1','=A1']])).toBe('#CYCLE!');
  expect(detectFormulaCycles({A:[['A1','="B1"']]})).toEqual([]);
  expect(detectFormulaCycles({A:[['=SUM(A1:B1)','1']]})).toEqual(['A!0:0']);
  const a=[['=B!A1']],b=[['=A!A1']];expect(evaluateFormula('=A1',a,{A:a,B:b})).toBe('#CYCLE!');
 });
 it('限制公式长度、嵌套、范围与递归深度',()=>{
  expect(evaluateFormula('='+ '1+'.repeat(6000)+'1',[])).toBe('#LIMIT!');
  expect(evaluateFormula('='+ '('.repeat(200)+'1'+')'.repeat(200),[])).toBe('#LIMIT!');
  expect(evaluateFormula('=SUM(A1:XFD1048576)',[])).toBe('#LIMIT!');
  const grid=Array.from({length:200},(_,r)=>[r===199?'1':'=A'+(r+2)]);
  expect(evaluateFormula('=A1',grid)).toBe('#LIMIT!');
 });
});
