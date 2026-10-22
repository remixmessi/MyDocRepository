import {it,expect} from 'vitest';
import {importableDrop,officeDrop,jsonImportKind} from '../packages/frontend/core/src/file-intake';
import {officeImportable,officeImportProblem,officeDelimited} from '../packages/common/nbstore/src/office';
const file=(name:string)=>new File(['x'],name);
it('识别可直接导入的办公文件，旧版格式与超限给出原因',()=>{
 for(const name of ['a.docx','b.XLSX','c.csv','d.tsv','e.pdf','f.pptx','g.dotm','h.potx'])expect(officeImportable(name)).toBe(true);
 for(const name of ['a.doc','b.xls','c.ppt','d.txt','e.md','docx','f.png'])expect(officeImportable(name)).toBe(false);
 expect(officeImportProblem('a.doc',1)).toContain('另存为 .docx');expect(officeImportProblem('a.png',1)).toContain('仅支持');
 expect(officeImportProblem('a.xlsx',20*1024*1024+1)).toContain('20 MB');expect(officeImportProblem('a.xlsx',10)).toBeUndefined();
 expect(officeDelimited('a.CSV')).toBe(true);expect(officeDelimited('a.xlsx')).toBe(false);
});
it('正文外拖入 Word / Excel / PPT 走办公导入，单独的 PDF 与正文内拖入仍作为附件',()=>{
 expect(officeDrop([file('报告.docx')],false)).toBe(true);expect(officeDrop([file('表.xlsx'),file('图.png')],false)).toBe(true);expect(officeDrop([file('数据.csv')],false)).toBe(true);
 expect(officeDrop([file('报告.docx')],true)).toBe(false);expect(officeDrop([file('资料.pdf')],false)).toBe(false);expect(officeDrop([file('资料.pdf'),file('报告.docx')],false)).toBe(true);
 expect(importableDrop([file('a.md'),file('b.docx')],false)).toBe(true);expect(importableDrop([file('b.docx')],false)).toBe(false);
});
it('按格式标记识别备份、草稿、模板 JSON，其他 JSON 不当作导入',()=>{
 expect(jsonImportKind(JSON.stringify({type:'mydoc-workspace',version:2}))).toBe('workspace-backup');
 expect(jsonImportKind(JSON.stringify({type:'mydoc-team-workspace',version:2}))).toBe('workspace-backup');
 expect(jsonImportKind(JSON.stringify({format:'mydoc-workspace-draft'}))).toBe('draft');expect(jsonImportKind(JSON.stringify({format:'mydoc-document-draft'}))).toBe('draft');
 expect(jsonImportKind(JSON.stringify({type:'mydoc-document-template'}))).toBe('document-template');expect(jsonImportKind(JSON.stringify({type:'mydoc-office-resource'}))).toBe('office-backup');
 expect(jsonImportKind(JSON.stringify({type:'mydoc-whiteboard-template'}))).toBe('whiteboard-template');
 for(const text of ['{"type":"other"}','[1]','null','不是 JSON','{"format":"mydoc-workspace"}'])expect(jsonImportKind(text)).toBeUndefined();
});
