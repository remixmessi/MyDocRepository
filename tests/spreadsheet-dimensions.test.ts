import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {mkdtempSync,rmSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {submitSheetOperation} from '../packages/frontend/apps/web/server/office-sheet-operation';
import {validateSpreadsheetUpdate,encodeXlsxWorkbook} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {exportXlsxText} from '../packages/frontend/apps/web/server/office-spreadsheet-export';
import {readOfficePackage,readOfficeXml} from '../packages/frontend/apps/web/server/office-package';
import {sheetColumnIds,sheetRowIdentity} from '../packages/common/nbstore/src/sheet-columns';
import {structureRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
import {sheetDimensions,sheetDimensionKey,validSheetDimension,excelWidthToPixels,pixelsToExcelWidth} from '../packages/common/nbstore/src/spreadsheet-dimensions';
import {sortedRows,sortedRowsBy} from '../packages/frontend/core/src/spreadsheet-commands';
const grid=(values:string[][])=>{const doc=new Y.Doc(),sheet=doc.getArray<Y.Array<string>>('sheet');for(const valuesRow of values){const row=new Y.Array<string>();row.push(valuesRow);sheet.push([row]);}return {doc,sheet};};

it('多列自定义排序：按第一关键字、并列时按第二关键字，空值始终在后，标题行不动；单列排序结果不变',()=>{
 const values=[['部门','销量'],['乙','30'],['甲','10'],['乙','5'],['','99'],['甲','40']],range={top:0,left:0,bottom:5,right:1};
 expect(sortedRowsBy(values,range,[{column:0,descending:false},{column:1,descending:true}],true).map(r=>values[r].join(''))).toEqual(['部门销量','甲40','甲10','乙30','乙5','99']);
 expect(sortedRowsBy(values,range,[{column:0,descending:true}],true).map(r=>values[r][0])).toEqual(['部门','乙','乙','甲','甲','']);
 expect(sortedRows(values,{top:1,left:0,bottom:5,right:1},1,false)).toEqual([3,2,1,5,4]);
 expect(excelWidthToPixels(pixelsToExcelWidth(140))).toBe(140);expect(validSheetDimension('column',{size:10})).toBe(false);expect(validSheetDimension('row',{hidden:true})).toBe(true);expect(validSheetDimension('row',{hidden:false})).toBe(false);
});

it('行高列宽与隐藏语义操作：按稳定标识保存，null 清除，删除行列与工作表时清理，非法值与已删除行列拒绝',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-sheet-dims-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),{doc,sheet}=grid([['a','b','c'],['d','e','f'],['g','h','i']]);try{
  cloud.database.append('cloud:space','book',Buffer.from(Y.encodeStateAsUpdate(doc)));const call=(input:Parameters<typeof submitSheetOperation>[4])=>{const result=submitSheetOperation(cloud.database,'space','book','actor',input);Y.applyUpdate(doc,result.update);return result;};
  const cols=()=>sheetColumnIds(sheet),rows=()=>sheet.toArray().map(sheetRowIdentity);
  const op={operationId:crypto.randomUUID(),type:'setDimensions' as const,base:'',axis:'column' as const,dimensions:[{key:cols()[1],size:180},{key:cols()[2],hidden:true as const}]};call(op);expect(call(op).repeated).toBe(true);
  call({operationId:crypto.randomUUID(),type:'setDimensions',base:'',axis:'row',dimensions:[{key:rows()[0],size:48,hidden:true}]});
  expect(sheetDimensions(doc,'first',sheet)).toEqual({rows:[{size:48,hidden:true},undefined,undefined],columns:[undefined,{size:180},{hidden:true}]});
  call({operationId:crypto.randomUUID(),type:'setDimensions',base:'',axis:'row',dimensions:[{key:rows()[0],hidden:null}]});expect(sheetDimensions(doc,'first',sheet).rows[0]).toEqual({size:48});
  // Insert a column before B: the width stays with the old B column, now C.
  call({operationId:crypto.randomUUID(),type:'insertColumn',column:1,base:structureRevision(doc)});expect(sheetDimensions(doc,'first',sheet).columns.map(value=>value?.size??value?.hidden??null)).toEqual([null,null,180,true]);
  const gone=cols()[2];call({operationId:crypto.randomUUID(),type:'deleteColumn',column:2,base:structureRevision(doc)});expect(doc.getMap('sheetDimensions').has(sheetDimensionKey('first','column',gone))).toBe(false);
  call({operationId:crypto.randomUUID(),type:'deleteRow',row:0,base:structureRevision(doc)});expect(doc.getMap('sheetDimensions').size).toBe(1);
  for(const bad of [{axis:'column',dimensions:[{key:cols()[0],size:5}]},{axis:'row',dimensions:[{key:cols()[0],size:30}]},{axis:'column',dimensions:[{key:cols()[0]}]},{axis:'column',dimensions:[{key:cols()[0],hidden:false}]},{axis:'diagonal',dimensions:[{key:cols()[0],size:40}]},{axis:'column',dimensions:[{key:cols()[0],size:40,extra:1}]}])
   expect(()=>submitSheetOperation(cloud.database,'space','book','actor',{operationId:crypto.randomUUID(),type:'setDimensions',base:'',...bad} as never)).toThrow('无效');
  expect(()=>submitSheetOperation(cloud.database,'space','book','actor',{operationId:crypto.randomUUID(),type:'setDimensions',base:'',axis:'column',dimensions:[{key:'column-first-99',size:40}]})).toThrow('已删除');
  const forged=new Y.Doc();Y.applyUpdate(forged,Y.encodeStateAsUpdate(doc));forged.getMap('sheetDimensions').set('first:c:x',{size:9999});expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(forged))).toThrow('行高列宽');forged.destroy();
 }finally{cloud.close();rmSync(dir,{recursive:true,force:true});doc.destroy();}
});

it('XLSX 往返：列宽与隐藏列写为 <cols>（相同设置合并区间），行高与隐藏行写在 <row>，重新导入还原',()=>{
 const {doc,sheet}=grid([['a','b','c','d'],['e','f','g','h'],['i','j','k','l']]),cols=sheetColumnIds(sheet),rows=sheet.toArray().map(sheetRowIdentity),map=doc.getMap('sheetDimensions');
 map.set(sheetDimensionKey('first','column',cols[1]),{size:160});map.set(sheetDimensionKey('first','column',cols[2]),{size:160});map.set(sheetDimensionKey('first','column',cols[3]),{hidden:true});map.set(sheetDimensionKey('first','row',rows[1]),{size:40});map.set(sheetDimensionKey('first','row',rows[2]),{hidden:true});
 try{
  const bytes=exportXlsxText(Y.encodeStateAsUpdate(doc)),xml=readOfficeXml(readOfficePackage(bytes)['xl/worksheets/sheet1.xml']);
  expect(xml).toContain(`<cols><col min="2" max="3" width="${pixelsToExcelWidth(160)}" customWidth="1"/><col min="4" max="4" width="${pixelsToExcelWidth(96)}" hidden="1"/></cols><sheetData>`);expect(xml).toContain('<row r="2" ht="30" customHeight="1">');expect(xml).toContain('<row r="3" hidden="1">');
  const back=encodeXlsxWorkbook(bytes);try{expect(sheetDimensions(back.doc,'first',back.doc.getArray('sheet'))).toEqual({rows:[undefined,{size:40},{hidden:true}],columns:[undefined,{size:160},{size:160},{hidden:true}]});}finally{back.doc.destroy();}
 }finally{doc.destroy();}
});
