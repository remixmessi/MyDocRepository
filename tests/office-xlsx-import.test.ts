import {it,expect} from 'vitest';
import {strToU8,zipSync} from 'fflate';
import * as Y from 'yjs';
import {encodeXlsxWorkbook,validateSpreadsheetUpdate} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {spreadsheetCellType} from '../packages/common/nbstore/src/spreadsheet-cell-type';
import {calculateFormulaGrid} from '../packages/common/nbstore/src/spreadsheet-formula';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';
import {validateOfficeBackupModel} from '../packages/frontend/apps/web/server/office-backup-model';
function workbook(cells:string,extras:Record<string,string>={}){return zipSync(Object.fromEntries(Object.entries({
 'xl/workbook.xml':`<x:workbook xmlns:x="urn:test" xmlns:r="urn:rel"><x:sheets><x:sheet name="选中 &amp; 首表" sheetId="7" r:id="rId7"/><x:sheet name="另一表" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`,
 'xl/_rels/workbook.xml.rels':`<Relationships><Relationship Id="rId1" Type="office/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId7" Type="office/worksheet" Target="/xl/worksheets/sheet7.xml"/><Relationship Id="s" Type="office/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
 'xl/worksheets/sheet1.xml':'<worksheet><sheetData><row r="1"><c r="A1"><v>999</v></c></row></sheetData></worksheet>',
 'xl/worksheets/sheet7.xml':`<worksheet><sheetData><row r='1'>${cells}</row></sheetData></worksheet>`,
 'xl/sharedStrings.xml':'<sst><si><r><t>001</t></r><rPh><t>不应拼入正文</t></rPh></si><si><t>=1+1</t></si></sst>',...extras,
 }).map(([key,value])=>[key,strToU8(value)])));}
it('按工作簿关系首表导入，保留普通公式并区分文本、数值和文本化的布尔/错误',()=>{
 const parsed=encodeXlsxWorkbook(workbook('<c r="A1" t="s"><v>0</v></c><c r="B1"><v>3</v></c><c r="C1"><f>SUM(B1:B1)</f><v>9999</v></c><c r="D1" t="s"><v>1</v></c><c r="E1" t="b"><v>1</v></c><c r="F1" t="e"><v>#DIV/0!</v></c><c r="G1" t="str"><f>"001"</f><v>wrong-cache</v></c>',{'xl/styles.xml':'<styleSheet/>'})),doc=parsed.doc;
 try{validateSpreadsheetUpdate(null,parsed.update);const values=doc.getArray<Y.Array<string>>('sheet').toArray().map(row=>row.toArray());expect(values[0]).toEqual(['001','3','=SUM(B1:B1)','=1+1','TRUE','#DIV/0!','="001"']);expect(values[0].map((_,c)=>spreadsheetCellType(doc,0,c))).toEqual(['text','auto','formula','text','text','text','formula']);expect(calculateFormulaGrid(values,(_s,r,c)=>spreadsheetCellType(doc,r,c))[0]).toEqual(['001','3','3','=1+1','TRUE','#DIV/0!','001']);expect(doc.getMap('sheetXlsx').toJSON()).toEqual({name:'选中 & 首表',sheetCount:2,hasStyles:true,textConversions:true});}finally{doc.destroy();}
});
it('支持命名空间、单引号属性、富文本、CDATA 和隐式地址，不重复解码',()=>{
 const parsed=encodeXlsxWorkbook(workbook('<c t="inlineStr"><is><r><t>&amp;quot; _x005F_x000A_</t></r><r><t><![CDATA[<中文😀>]]></t></r></is></c><c><f><![CDATA[1+1]]></f><v>123</v></c>'));
 try{expect(parsed.doc.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(['&quot; _x000A_<中文😀>','=1+1']);}finally{parsed.doc.destroy();}
});
it('共享/数组公式、错误索引、重复或越界地址和未知类型明确拒绝',()=>{
 for(const value of ['<c r="A1"><f t="shared" si="0">A2</f><v>3</v></c>','<c r="A1"><f t="array" ref="A1:A2">A2</f></c>','<c r="A1"><f t="shared" si="0"/><v>3</v></c>'])expect(()=>encodeXlsxWorkbook(workbook(value))).toThrow('暂不支持');
 for(const value of ['<c r="A1" t="s"><v>999</v></c>','<c r="A1"><v>1</v></c><c r="A1"><v>2</v></c>','<c r="XFE1"><v>1</v></c>','<c r="A10001"><v>1</v></c>','<c r="A1" t="d"><v>2026-10-18</v></c>','<c r="A1"><v>1e999</v></c>','<c r="A1" t="inlineStr"><is><t>&bad</t></is></c>'])expect(()=>encodeXlsxWorkbook(workbook(value))).toThrow();
 expect(()=>encodeXlsxWorkbook(workbook('',{'xl/worksheets/sheet7.xml':'<worksheet><sheetData><row r="1"><c r="AZ1"><v>1</v></c></row><row r="10000"/></sheetData></worksheet>'}))).toThrow('20 万');
});
it('关系缺失、外部首表及路径越界拒绝，隔离模型和原生备份保留公式与来源说明',async()=>{
 for(const relation of ['<Relationships/>','<Relationships><Relationship Id="rId7" Type="office/worksheet" Target="https://example.com/sheet.xml" TargetMode="External"/></Relationships>','<Relationships><Relationship Id="rId7" Type="office/worksheet" Target="../../evil.xml"/></Relationships>'])expect(()=>encodeXlsxWorkbook(workbook('',{'xl/_rels/workbook.xml.rels':relation}))).toThrow();
 const pool=new OfficeParserPool(),doc=new Y.Doc();try{const bytes=workbook('<c r="A1"><f>1+1</f><v>999</v></c>'),update=await pool.parse('xlsx',bytes);Y.applyUpdate(doc,validateOfficeBackupModel('spreadsheet',update));expect(doc.getArray<Y.Array<string>>('sheet').get(0).get(0)).toBe('=1+1');expect(spreadsheetCellType(doc,0,0)).toBe('formula');expect(doc.getMap('sheetXlsx').get('sheetCount')).toBe(2);doc.getMap('sheetXlsx').set('foreign',true);expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('导入范围');}finally{pool.close();doc.destroy();}
});
it('不支持或超出范围的条件格式与图表计入导入范围信息，计数只在大于 0 时写入',()=>{
 const cf='<conditionalFormatting sqref="A1:A2"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule><cfRule type="expression" priority="2"><formula>A1&gt;1</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="bad"><cfRule type="colorScale" priority="3"><colorScale><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>';
 const parsed=encodeXlsxWorkbook(workbook('',{
  'xl/worksheets/sheet7.xml':`<worksheet xmlns:r="urn:rel"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData>${cf}<drawing r:id="rIdD"/></worksheet>`,
  'xl/worksheets/_rels/sheet7.xml.rels':'<Relationships><Relationship Id="rIdD" Type="office/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
  'xl/drawings/drawing1.xml':'<wsDr xmlns:r="urn:rel"><twoCellAnchor><from><col>2</col><row>0</row></from><to><col>6</col><row>10</row></to><graphicFrame><graphic><graphicData><chart r:id="rIdMissing"/></graphicData></graphic></graphicFrame></twoCellAnchor></wsDr>',
 }));
 try{validateSpreadsheetUpdate(null,parsed.update);expect(parsed.doc.getMap('sheetXlsx').toJSON()).toEqual({name:'选中 & 首表',sheetCount:2,hasStyles:false,textConversions:false,skippedRules:2,skippedCharts:1});expect(parsed.doc.getMap('sheetObjects').size).toBe(1);}finally{parsed.doc.destroy();}
 const clean=encodeXlsxWorkbook(workbook('<c r="A1"><v>1</v></c>'));try{expect(Object.keys(clean.doc.getMap('sheetXlsx').toJSON())).not.toContain('skippedRules');
  for(const [key,value]of [['skippedRules',0],['skippedCharts',1.5],['skippedCharts','1']] as const){const doc=new Y.Doc();Y.applyUpdate(doc,clean.update);doc.getMap('sheetXlsx').set(key,value);expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('导入范围');doc.destroy();}
 }finally{clean.doc.destroy();}
});
