import {it,expect} from 'vitest';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {wordWritingEnd,splitWordParagraph,mergeWordParagraph,adjacentWordParagraph} from '../packages/frontend/core/src/office-word-flow';
import type {WordBlock} from '../packages/common/nbstore/src/office-word-table';

it('空白正文直接可写，重复进入不会新增空段；表格/图片之后可继续书写',()=>{
 const doc=new Y.Doc(),blocks=doc.getArray<WordBlock>('paragraphs');
 const first=wordWritingEnd(doc);expect(blocks.toArray()).toEqual([first]);expect(wordWritingEnd(doc)).toBe(first);
 const image=new Y.Map();image.set('kind','image');blocks.push([image]);const tail=wordWritingEnd(doc);expect(tail).not.toBe(first);expect(blocks.toArray()).toEqual([first,image,tail]);expect(wordWritingEnd(doc)).toBe(tail);doc.destroy();
});
it('回车按选区拆分正文并保留文字格式；新段不继承段前分页和末尾标题',()=>{
 const doc=new Y.Doc(),text=wordWritingEnd(doc);text.applyDelta([{insert:'甲😀',attributes:{bold:true}},{insert:'乙丙',attributes:{italic:true,color:'#FF0000'}}]);text.setAttribute('pageBreakBefore',true);text.setAttribute('align','center');text.setAttribute('heading',2);
 const next=splitWordParagraph(text,1,3);expect(text.toString()).toBe('甲');expect(next.toString()).toBe('乙丙');expect(next.toDelta()).toEqual([{insert:'乙丙',attributes:{italic:true,color:'#FF0000'}}]);expect(next.getAttribute('align')).toBe('center');expect(next.getAttribute('pageBreakBefore')).toBeUndefined();expect(next.getAttribute('heading')).toBe(2);
 const normal=splitWordParagraph(next,next.length);expect(normal.getAttribute('heading')).toBeUndefined();expect(normal.toString()).toBe('');doc.destroy();
});
it('段首退格和段尾 Delete 保留格式，合并可整体撤销与重做',()=>{
 const doc=new Y.Doc(),first=wordWritingEnd(doc);first.insert(0,'甲',{bold:true});const second=splitWordParagraph(first,first.length);second.insert(0,'乙',{italic:true});
 const undo=new Y.UndoManager(doc.getArray('paragraphs'),{trackedOrigins:new Set([LOCAL_ORIGIN])});const merged=mergeWordParagraph(second,'backward');expect(merged).toEqual({text:first,offset:1});expect(first.toDelta()).toEqual([{insert:'甲',attributes:{bold:true}},{insert:'乙',attributes:{italic:true}}]);expect(doc.getArray('paragraphs')).toHaveLength(1);undo.undo();expect(doc.getArray<Y.Text>('paragraphs').toArray().map(text=>text.toString())).toEqual(['甲','乙']);undo.redo();expect(first.toString()).toBe('甲乙');undo.destroy();doc.destroy();
});
it('合并不跨越表格/图片，失效对象及 emoji 中间位置不改写正文',()=>{
 const doc=new Y.Doc(),first=wordWritingEnd(doc);first.insert(0,'😀甲');const blocks=doc.getArray<WordBlock>('paragraphs'),image=new Y.Map();image.set('kind','image');blocks.push([image]);const last=wordWritingEnd(doc);expect(mergeWordParagraph(last,'backward')).toBeUndefined();expect(adjacentWordParagraph(first,'forward')).toBeUndefined();expect(()=>splitWordParagraph(first,1)).toThrow('文字位置');expect(first.toString()).toBe('😀甲');blocks.delete(0,1);expect(()=>splitWordParagraph(first,0)).toThrow('正文位置');doc.destroy();
});
it('已接收的远端内容在换段与合并后同步一致',()=>{
 const a=new Y.Doc(),b=new Y.Doc(),text=wordWritingEnd(a);text.insert(0,'甲乙');Y.applyUpdate(b,Y.encodeStateAsUpdate(a));b.getArray<Y.Text>('paragraphs').get(0).insert(1,'远端');Y.applyUpdate(a,Y.encodeStateAsUpdate(b));const next=splitWordParagraph(text,1);expect(next.toString()).toBe('远端乙');mergeWordParagraph(text,'forward');Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(b.getArray<Y.Text>('paragraphs').toArray().map(text=>text.toString())).toEqual(['甲远端乙']);a.destroy();b.destroy();
});

it('Word 表格列操作按活动单元格定位，保留文字格式、可撤销，并支持原生往返',async()=>{
 const {appendWordTable,insertWordTableColumn,deleteWordTableColumn}=await import('../packages/frontend/core/src/office-word-table');const {wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {validateWordUpdate,encodeDocxWord}=await import('../packages/frontend/apps/web/server/office-word');const {exportDocxWord}=await import('../packages/frontend/apps/web/server/office-word-export');
 const doc=new Y.Doc(),table=appendWordTable(doc,2,2),rows=wordTableRows(table),anchor=rows.get(1).get(1);anchor.insert(0,'保留😀',{bold:true});const undo=new Y.UndoManager(doc.getArray('paragraphs'),{trackedOrigins:new Set([LOCAL_ORIGIN])});
 insertWordTableColumn(table,anchor,'before');expect(rows.get(1).get(2)).toBe(anchor);expect(table.get('columns')).toBe(3);undo.stopCapturing();insertWordTableColumn(table,anchor,'after');expect(rows.get(1)).toHaveLength(4);expect(anchor.toDelta()).toEqual([{insert:'保留😀',attributes:{bold:true}}]);
 undo.undo();expect(rows.get(1)).toHaveLength(3);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();const parsed=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(doc)));expect(wordTableRows(parsed.doc.getArray<Y.Map<unknown>>('paragraphs').get(0)).get(1).get(2).toString()).toBe('保留😀');parsed.doc.destroy();
 deleteWordTableColumn(table,anchor);expect(table.get('columns')).toBe(2);expect(()=>insertWordTableColumn(table,anchor,'after')).toThrow('已删除');deleteWordTableColumn(table,rows.get(0).get(0));expect(()=>deleteWordTableColumn(table,rows.get(0).get(0))).toThrow('至少保留');undo.destroy();doc.destroy();
});

it('跨段选区格式、替换与本人撤销保留尾部格式，远端插入后相对定位仍有效',async()=>{
 const {captureWordFlowSelection,formatWordFlow,replaceWordFlow,wordFlowPlainText}=await import('../packages/frontend/core/src/word-flow-selection');
 const doc=new Y.Doc(),remote=new Y.Doc();try{
  const blocks=doc.getArray<WordBlock>('paragraphs'),first=new Y.Text('甲乙😀'),middle=new Y.Text('中间'),last=new Y.Text('丙丁戊');blocks.push([first,middle,last]);last.format(1,2,{italic:true});
  const selection=captureWordFlowSelection(first,1,last,1);Y.applyUpdate(remote,Y.encodeStateAsUpdate(doc));(remote.getArray('paragraphs').get(0) as Y.Text).insert(0,'远端');Y.applyUpdate(doc,Y.encodeStateAsUpdate(remote));expect(wordFlowPlainText(doc,selection)).toBe('乙😀\n中间\n丙');formatWordFlow(doc,selection,'bold');expect(first.toDelta()).toEqual([{insert:'远端甲'},{insert:'乙😀',attributes:{bold:true}}]);expect(middle.toDelta()).toEqual([{insert:'中间',attributes:{bold:true}}]);
  formatWordFlow(doc,selection,'fontSize',18);const undo=new Y.UndoManager(blocks,{trackedOrigins:new Set([LOCAL_ORIGIN])});try{const result=replaceWordFlow(doc,selection,'替换\n新增');expect(blocks.toArray().map(text=>String(text))).toEqual(['远端甲替换','新增丁戊']);expect(result.offset).toBe(2);expect(result.text.toDelta().at(-1)).toEqual({insert:'丁戊',attributes:{italic:true}});undo.undo();expect(blocks.toArray().map(text=>String(text))).toEqual(['远端甲乙😀','中间','丙丁戊']);}finally{undo.destroy();}
 }finally{doc.destroy();remote.destroy();}
});
it('跨段选区拒绝跨越表格、失效端点和代理字符中间位置，不出现部分修改',async()=>{
 const {captureWordFlowSelection,replaceWordFlow}=await import('../packages/frontend/core/src/word-flow-selection');const doc=new Y.Doc();try{const blocks=doc.getArray<WordBlock>('paragraphs'),first=new Y.Text('😀甲'),last=new Y.Text('乙');const table=new Y.Map<unknown>();blocks.push([first,table,last]);expect(()=>captureWordFlowSelection(first,1,last,0)).toThrow('文字位置');const selection=captureWordFlowSelection(first,0,last,1),before=Y.encodeStateAsUpdate(doc);expect(()=>replaceWordFlow(doc,selection,'x')).toThrow('跨越表格');expect(Y.encodeStateAsUpdate(doc)).toEqual(before);blocks.delete(2,1);expect(()=>replaceWordFlow(doc,selection,'x')).toThrow('已删除');}finally{doc.destroy();}
});

it('自动分页按正文高度换页：段前分页、整块图片/表格行、段落按行续页，空文档仅一页',async()=>{
 const {paginateWordBlocks}=await import('../packages/frontend/core/src/use-word-pagination');
 expect(paginateWordBlocks([],900)).toEqual([]);
 const lines=(n:number)=>Array.from({length:n},(_,i)=>i*20);
 const blocks=[{top:0,height:400,lines:lines(20),breakBefore:false,whole:false},{top:400,height:300,lines:[],breakBefore:false,whole:true},{top:700,height:400,lines:lines(20),breakBefore:false,whole:false},{top:1100,height:20,lines:[0],breakBefore:true,whole:false},{top:1120,height:900,lines:[],breakBefore:false,whole:true}];
 // Page body 900px: the image fits on page 1; the next paragraph breaks at its line 10 (y=900); the manual break starts page 3; an oversized image starts page 4 and simply overflows it.
 expect(paginateWordBlocks(blocks,900)).toEqual([{top:900,page:2},{top:1100,page:3},{top:1120,page:4}]);
 expect(paginateWordBlocks([{top:0,height:1900,lines:lines(95),breakBefore:false,whole:false}],900)).toEqual([{top:900,page:2},{top:1800,page:3}]);
});
