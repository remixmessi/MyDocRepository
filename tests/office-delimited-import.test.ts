import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {encodeOfficeText,decodeOfficeText} from '../packages/common/nbstore/src/office';
import {prepareDelimitedImport,validDelimitedImportOptions,validDelimitedImportSettings} from '../packages/common/nbstore/src/office-delimited-import';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';
import {validateSpreadsheetUpdate} from '../packages/frontend/apps/web/server/office-spreadsheet';
import {validateOfficeBackupModel} from '../packages/frontend/apps/web/server/office-backup-model';
const csv={delimiter:',' as const,quoteMode:'csv' as const};
it('严格 CSV 解析保留首行、前导零、长编号、公式原文、Unicode、字段内换行与空末列',()=>{
 const text='标题,编号,公式,尾列\r\n"中文😀\r\n第二行",00123,=1+1,\r\n"a""b",12345678901234567890,"",\r\n';
 const parsed=prepareDelimitedImport(encodeOfficeText(text,'utf-8',true),csv);expect(parsed.rows).toEqual([['标题','编号','公式','尾列'],['中文😀\r\n第二行','00123','=1+1',''],['a"b','12345678901234567890','','']]);expect(parsed.settings).toEqual({encoding:'utf-8',bom:true,...csv});
 expect(prepareDelimitedImport(new Uint8Array(),csv).rows).toEqual([]);expect(prepareDelimitedImport(new TextEncoder().encode('""'),csv).rows).toEqual([['']]);expect(prepareDelimitedImport(new TextEncoder().encode(','),csv).rows).toEqual([['','']]);
});
it('TSV 显式无引号模式保留引号字面值，严格模式拒绝歧义',()=>{
 const bytes=new TextEncoder().encode('a"b\t"中文"\t001\n');expect(()=>prepareDelimitedImport(bytes,{delimiter:'\t',quoteMode:'csv'})).toThrow('引号');expect(prepareDelimitedImport(bytes,{delimiter:'\t',quoteMode:'none'}).rows).toEqual([['a"b','"中文"','001']]);
 for(const text of ['"unclosed','a"b,c','"a"x,b'])expect(()=>prepareDelimitedImport(new TextEncoder().encode(text),csv)).toThrow('引号');
 expect(validDelimitedImportOptions({delimiter:',',quoteMode:'none'})).toBe(false);expect(validDelimitedImportOptions({delimiter:['\t'],quoteMode:'csv'})).toBe(false);expect(validDelimitedImportOptions({...csv,foreign:true})).toBe(false);
});
it('UTF 编码、BOM 冲突、非法编码和所有解析配额在创建模型前校验',()=>{
 for(const encoding of ['utf-8','utf-16le','utf-16be'] as const)for(const bom of [false,true]){const parsed=prepareDelimitedImport(encodeOfficeText('中文😀,001\r\n',encoding,bom),{...csv,...(!bom?{encoding}:{})});expect(parsed.rows).toEqual([['中文😀','001']]);expect(parsed.settings).toEqual({...csv,encoding,bom});}
 expect(()=>decodeOfficeText(encodeOfficeText('a','utf-8',true),'utf-16le')).toThrow('BOM');expect(()=>prepareDelimitedImport(new Uint8Array([255]),csv)).toThrow();expect(()=>prepareDelimitedImport(new Uint8Array([0,216]),{...csv,encoding:'utf-16le'})).toThrow();
 for(const [text,message]of [['a'.repeat(10001),'1 万字符'],['\n'.repeat(10001),'1 万'],[','.repeat(200000),'20 万']])expect(()=>prepareDelimitedImport(new TextEncoder().encode(text),csv)).toThrow(message);
 expect(()=>prepareDelimitedImport(new Uint8Array(20*1024*1024+1),csv)).toThrow('20 MB');
});
it('隔离进程导入保存设置，原生备份保留设置且非法设置拒绝',async()=>{
 const pool=new OfficeParserPool(),doc=new Y.Doc();try{const bytes=encodeOfficeText('"标题"\t001\n','utf-16be',false),update=await pool.parseDelimited(bytes,{encoding:'utf-16be',delimiter:'\t',quoteMode:'none'});Y.applyUpdate(doc,update);expect(doc.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(['"标题"','001']);const settings=doc.getMap('sheetImport').toJSON();expect(validDelimitedImportSettings(settings)).toBe(true);const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('spreadsheet',update));expect(restored.getMap('sheetImport').toJSON()).toEqual(settings);}finally{restored.destroy();}doc.getMap('sheetImport').set('foreign',true);expect(()=>validateSpreadsheetUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('导入设置');await expect(pool.parseDelimited(bytes,{...csv,quoteMode:'invalid'} as never)).rejects.toMatchObject({status:400});}finally{pool.close();doc.destroy();}
});

it('剪贴板解析保留字面 BOM、引号换行、空末字段，拒绝歧义引号',async()=>{
 const {parseDelimitedGrid}=await import('../packages/common/nbstore/src/office-delimited-import');
 expect(parseDelimitedGrid('\ufeff001\t=1+1\r\n"中文\n末尾"\t\r\n','\t')).toEqual([['\ufeff001','=1+1'],['中文\n末尾','']]);
 expect(parseDelimitedGrid('""','\t')).toEqual([['']]);expect(()=>parseDelimitedGrid('"a"b','\t')).toThrow('结束引号');expect(()=>parseDelimitedGrid('a"b','\t')).toThrow('转义');
});
