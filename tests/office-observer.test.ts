import {it,expect,vi} from 'vitest';
import {Y} from '@mydoc/crdt';
import {observeOfficeDocument} from '../packages/frontend/core/src/office-observer';
import {applyOfficeTextEdit} from '../packages/frontend/core/src/office-text';

it('嵌套段落和对象的远端更新触发通知，清理后停止通知',()=>{
 const a=new Y.Doc(),b=new Y.Doc();
 const paragraph=new Y.Text('原文');a.getArray('paragraphs').push([paragraph]);
 const object=new Y.Map();object.set('content','批注');a.getArray('objects').push([object]);
 Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 const notify=vi.fn(),stop=observeOfficeDocument(b,notify);
 paragraph.insert(0,'远端');object.set('content','新批注');
 Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(notify).toHaveBeenCalledTimes(1);
 expect((b.getArray('paragraphs').get(0) as Y.Text).toString()).toBe('远端原文');
 stop();paragraph.insert(0,'再次');Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(notify).toHaveBeenCalledTimes(1);
 a.destroy();b.destroy();
});

it('DOCX 局部字符编辑合并两个客户端的修改，不重写整段',()=>{
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('原文')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 const pa=a.getArray<Y.Text>('paragraphs').get(0),pb=b.getArray<Y.Text>('paragraphs').get(0);
 applyOfficeTextEdit(pa,'原文','甲原文');applyOfficeTextEdit(pb,'原文','原文乙');
 const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
 expect(pa.toString()).toBe('甲原文乙');expect(pb.toString()).toBe(pa.toString());a.destroy();b.destroy();
});

it('DOCX 组合输入使用嵌套段落快照保留远端字符',()=>{
 const live=new Y.Doc(),snapshot=new Y.Doc();live.getArray('paragraphs').push([new Y.Text('原文')]);
 Y.applyUpdate(snapshot,Y.encodeStateAsUpdate(live));const text=live.getArray<Y.Text>('paragraphs').get(0);
 const position=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromTypeIndex(text,0),snapshot)!;
 const reference=position.type as Y.Text;text.insert(0,'远端');
 applyOfficeTextEdit(text,reference.toString(),'原中文文',reference);expect(text.toString()).toBe('远端原中文文');
 live.destroy();snapshot.destroy();
});

it('Word 样式选区随远端前插移动，字符编辑保留样式且并发修改收敛',async()=>{
 const {captureWordSelection,toggleWordMark}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('原文😀')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const ta=a.getArray<Y.Text>('paragraphs').get(0),tb=b.getArray<Y.Text>('paragraphs').get(0);
  const selection=captureWordSelection(ta,0,2)!;
  tb.insert(0,'远端');Y.applyUpdate(a,Y.encodeStateAsUpdate(b));toggleWordMark(ta,selection,'bold');
  expect(ta.toDelta()).toEqual([{insert:'远端'},{insert:'原文',attributes:{bold:true}},{insert:'😀'}]);
  Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
  applyOfficeTextEdit(ta,ta.toString(),'远端原中文😀');
  toggleWordMark(tb,captureWordSelection(tb,4,6)!,'italic');
  const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
  expect(ta.toString()).toBe('远端原中文😀');expect(ta.toDelta()).toEqual(tb.toDelta());expect(ta.toDelta().some((part:{attributes?:Record<string,boolean>})=>part.attributes?.bold===true)).toBe(true);
 }finally{a.destroy();b.destroy();}
});
it('Word 切换整段混合样式、已删除选区和 emoji 边界有确定结果',async()=>{
 const {captureWordSelection,toggleWordMark}=await import('../packages/frontend/core/src/office-word-format');
 const doc=new Y.Doc();doc.getArray('paragraphs').push([new Y.Text('a😀b')]);
 try{
  const text=doc.getArray<Y.Text>('paragraphs').get(0);text.format(0,1,{underline:true});
  const all=captureWordSelection(text,0,4)!;toggleWordMark(text,all,'underline');expect(text.toDelta()).toEqual([{insert:'a😀b',attributes:{underline:true}}]);
  toggleWordMark(text,all,'underline');expect(text.toDelta()).toEqual([{insert:'a😀b',attributes:{underline:false}}]);
  toggleWordMark(text,captureWordSelection(text,2,3)!,'bold');expect(text.toDelta().find((part:{attributes?:Record<string,boolean>})=>part.attributes?.bold)?.insert).toBe('😀');
  const removed=captureWordSelection(text,1,3)!;text.delete(1,2);expect(()=>toggleWordMark(text,removed,'italic')).toThrow('已删除');
 }finally{doc.destroy();}
});

it('标题切换与远端字符、样式修改收敛，取消标题不丢文字，删除段落不能设置标题',async()=>{
 const {setWordHeading}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('原文')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const ta=a.getArray<Y.Text>('paragraphs').get(0),tb=b.getArray<Y.Text>('paragraphs').get(0);
  setWordHeading(ta,2);tb.insert(0,'远端');tb.format(0,2,{italic:true});
  const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
  expect(tb.getAttribute('heading')).toBe(2);expect(ta.toDelta()).toEqual(tb.toDelta());expect(ta.toString()).toBe('远端原文');
  setWordHeading(tb,0);Y.applyUpdate(a,Y.encodeStateAsUpdate(b));expect(ta.getAttribute('heading')).toBe(0);expect(ta.toString()).toBe('远端原文');
  a.getArray('paragraphs').delete(0,1);expect(()=>setWordHeading(ta,1)).toThrow('删除');
 }finally{a.destroy();b.destroy();}
});

it('列表分组连续编号，独立组重启，移除列表和并发文字不丢内容',async()=>{
 const {setWordList,wordListMarkers}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push(['一','二','正文','三'].map(value=>new Y.Text(value)));
 try{
  const p=a.getArray<Y.Text>('paragraphs');setWordList(p.get(0),'decimal');setWordList(p.get(1),'decimal');setWordList(p.get(3),'decimal');expect(wordListMarkers(p.toArray())).toEqual(['1.','2.','','1.']);
  Y.applyUpdate(b,Y.encodeStateAsUpdate(a));setWordList(p.get(0),'bullet');b.getArray<Y.Text>('paragraphs').get(0).insert(0,'远端');const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);expect(p.get(0).toString()).toBe('远端一');expect(p.get(0).getAttribute('list')).toEqual(b.getArray<Y.Text>('paragraphs').get(0).getAttribute('list'));
  setWordList(p.get(1),'none');expect(wordListMarkers(p.toArray())).toEqual(['•','','','1.']);expect(p.get(1).toString()).toBe('二');
 }finally{a.destroy();b.destroy();}
});

it('链接设置跟随相对选区，保留并发文字与格式，移除不丢文字',async()=>{
 const {captureWordSelection,setWordLink}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('链接😀')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const ta=a.getArray<Y.Text>('paragraphs').get(0),tb=b.getArray<Y.Text>('paragraphs').get(0),selection=captureWordSelection(ta,0,2)!;
  tb.insert(0,'远端');tb.format(2,2,{bold:true});Y.applyUpdate(a,Y.encodeStateAsUpdate(b));setWordLink(ta,selection,'https://example.com/?a=1&b=2');
  expect(ta.toDelta()).toEqual([{insert:'远端'},{insert:'链接',attributes:{bold:true,href:'https://example.com/?a=1&b=2'}},{insert:'😀'}]);
  Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(tb.toDelta()).toEqual(ta.toDelta());setWordLink(ta,selection);expect(ta.toString()).toBe('远端链接😀');expect(ta.toDelta()).toEqual([{insert:'远端'},{insert:'链接',attributes:{bold:true}},{insert:'😀'}]);
  const before=ta.toDelta();expect(()=>setWordLink(ta,selection,'javascript:alert(1)')).toThrow('HTTP/HTTPS');expect(ta.toDelta()).toEqual(before);
  ta.delete(2,2);expect(()=>setWordLink(ta,selection,'https://example.com/')).toThrow('已删除');
 }finally{a.destroy();b.destroy();}
});

it('对齐、缩进与并发文字修改独立收敛，非法布局和已删除段落拒绝修改',async()=>{
 const {setWordAlignment,setWordLeftIndent}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('原文')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const ta=a.getArray<Y.Text>('paragraphs').get(0),tb=b.getArray<Y.Text>('paragraphs').get(0);setWordAlignment(ta,'center');setWordLeftIndent(tb,720);tb.insert(0,'远端');const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
  expect(ta.getAttributes()).toEqual({align:'center',leftIndent:720});expect(tb.getAttributes()).toEqual(ta.getAttributes());expect(ta.toString()).toBe('远端原文');expect(()=>setWordLeftIndent(ta,-1)).toThrow('缩进');a.getArray('paragraphs').delete(0,1);expect(()=>setWordAlignment(ta,'right')).toThrow('删除');
 }finally{a.destroy();b.destroy();}
});

it('右缩进与特殊缩进分别合并，模式竞争原子收敛，不改正文或其他属性',async()=>{
 const {setWordRightIndent,setWordFirstLineIndent}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('正文')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const ta=a.getArray<Y.Text>('paragraphs').get(0),tb=b.getArray<Y.Text>('paragraphs').get(0);setWordRightIndent(ta,720);setWordFirstLineIndent(ta,360);setWordFirstLineIndent(tb,-720);tb.insert(0,'远端');const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
  expect(ta.getAttributes()).toEqual(tb.getAttributes());expect(ta.getAttribute('rightIndent')).toBe(720);expect([360,-720]).toContain(ta.getAttribute('firstLineIndent'));expect(ta.toString()).toBe('远端正文');setWordFirstLineIndent(ta,0);expect(ta.getAttribute('rightIndent')).toBe(720);expect(ta.getAttribute('firstLineIndent')).toBe(0);
  expect(()=>setWordRightIndent(ta,-1)).toThrow('缩进');a.getArray('paragraphs').delete(0,1);expect(()=>setWordFirstLineIndent(ta,360)).toThrow('删除');
 }finally{a.destroy();b.destroy();}
});

it('页面设置和段前分页与文字独立合并，旧设置保存被提示，页面竞争原子收敛',async()=>{
 const {getWordPage,setWordPage,setWordPageBreak}=await import('../packages/frontend/core/src/office-word-page');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('正文')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const base=getWordPage(a),wide={...base,width:base.height,height:base.width},changed={...base,top:720};setWordPage(a,wide,base);setWordPage(b,changed,base);b.getArray<Y.Text>('paragraphs').get(0).insert(0,'远端');setWordPageBreak(a.getArray<Y.Text>('paragraphs').get(0),true);const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
  expect(getWordPage(a)).toEqual(getWordPage(b));expect([wide,changed]).toContainEqual(getWordPage(a));expect(a.getArray<Y.Text>('paragraphs').get(0).toString()).toBe('远端正文');expect(b.getArray<Y.Text>('paragraphs').get(0).getAttribute('pageBreakBefore')).toBe(true);
  setWordPage(a,wide);expect(()=>setWordPage(a,changed,base)).toThrow('其他成员');expect(getWordPage(a)).toEqual(wide);expect(()=>setWordPage(a,{...wide,left:wide.width})).toThrow('正文区域');
 }finally{a.destroy();b.destroy();}
});

it('字号与颜色在相对选区上独立合并，混合检测与恢复继承不删除其他属性',async()=>{
 const {captureWordSelection,setWordFont,wordSelectionFont}=await import('../packages/frontend/core/src/office-word-format');
 const a=new Y.Doc(),b=new Y.Doc();a.getArray('paragraphs').push([new Y.Text('文字😀')]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{
  const ta=a.getArray<Y.Text>('paragraphs').get(0),tb=b.getArray<Y.Text>('paragraphs').get(0),selection=captureWordSelection(ta,0,2)!;tb.insert(0,'远端');Y.applyUpdate(a,Y.encodeStateAsUpdate(b));setWordFont(ta,selection,'fontSize',13.5);setWordFont(tb,captureWordSelection(tb,2,4)!,'color','#123abc');const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
  expect(ta.toDelta()).toEqual(tb.toDelta());expect(wordSelectionFont(ta,2,4,'fontSize')).toBe(13.5);expect(wordSelectionFont(ta,0,4,'color')).toBe('mixed');setWordFont(ta,selection,'fontSize',undefined);expect(wordSelectionFont(ta,2,4,'fontSize')).toBeUndefined();expect(wordSelectionFont(ta,2,4,'color')).toBe('#123ABC');expect(()=>setWordFont(ta,selection,'color','url(x)')).toThrow('无效');ta.delete(2,2);expect(()=>setWordFont(ta,selection,'fontSize',12)).toThrow('已删除');
 }finally{a.destroy();b.destroy();}
});

it('Word 表格同格字符合并与并发新增行收敛，固定列结构保持完整',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {appendWordTableRow}=await import('../packages/frontend/core/src/office-word-table');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 const a=new Y.Doc(),b=new Y.Doc(),ta=createWordTable(2,2);a.getArray('paragraphs').push([ta]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{const tb=b.getArray<Y.Map<unknown>>('paragraphs').get(0);wordTableRows(ta).get(0).get(0).insert(0,'甲');wordTableRows(tb).get(0).get(0).insert(0,'乙');appendWordTableRow(ta);appendWordTableRow(tb);const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);expect(wordTableRows(ta).length).toBe(4);expect(wordTableRows(ta).get(0).get(0).toString()).toBe(wordTableRows(tb).get(0).get(0).toString());expect([...wordTableRows(ta).get(0).get(0).toString()].sort()).toEqual(['乙','甲'].sort());expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(a))).not.toThrow();}finally{a.destroy();b.destroy();}
});
it('删除表格行不误删远端新增行，删至零行可恢复，已删除对象不能复活',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {appendWordTableRow,deleteWordTableRow}=await import('../packages/frontend/core/src/office-word-table');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 const a=new Y.Doc(),b=new Y.Doc(),ta=createWordTable(2,2);a.getArray('paragraphs').push([ta]);Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
 try{const tb=b.getArray<Y.Map<unknown>>('paragraphs').get(0),deleted=wordTableRows(ta).get(0);deleteWordTableRow(ta,deleted);wordTableRows(tb).get(0).get(0).insert(0,'删除期间输入');appendWordTableRow(tb);Y.applyUpdate(a,Y.encodeStateAsUpdate(b));Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(wordTableRows(ta).length).toBe(2);expect(()=>deleteWordTableRow(ta,deleted)).toThrow('已删除');deleteWordTableRow(ta,wordTableRows(ta).get(0));deleteWordTableRow(tb,wordTableRows(tb).get(1));const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);expect(wordTableRows(ta).length).toBe(0);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(a))).not.toThrow();appendWordTableRow(ta);expect(wordTableRows(ta).length).toBe(1);a.getArray('paragraphs').delete(0,1);expect(()=>appendWordTableRow(ta)).toThrow('已删除');}finally{a.destroy();b.destroy();}
});

it('表格组合输入通过嵌套单元格快照合并远端文字，删除行后拒绝旧锚点',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {deleteWordTableRow}=await import('../packages/frontend/core/src/office-word-table');
 const live=new Y.Doc(),snapshot=new Y.Doc();
 try{
  const table=createWordTable(2,2);live.getArray('paragraphs').push([table]);const row=wordTableRows(table).get(0),text=row.get(0);text.insert(0,'原文');Y.applyUpdate(snapshot,Y.encodeStateAsUpdate(live));
  const reference=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromTypeIndex(text,0),snapshot)!.type as Y.Text;
  text.insert(0,'远端');applyOfficeTextEdit(text,reference.toString(),'原中文文',reference);expect(text.toString()).toBe('远端原中文文');
  deleteWordTableRow(table,row);expect(()=>applyOfficeTextEdit(text,reference.toString(),'原另文',reference)).toThrow('失效');
 }finally{live.destroy();snapshot.destroy();}
});

it('表格后的正文可设置列表，不把表格当作相邻段落',async()=>{
 const {createWordTable}=await import('../packages/common/nbstore/src/office-word-table');const {setWordList}=await import('../packages/frontend/core/src/office-word-format');const doc=new Y.Doc();
 try{const text=new Y.Text('表后正文');doc.getArray('paragraphs').push([createWordTable(2,2),text]);expect(()=>setWordList(text,'decimal')).not.toThrow();expect(text.getAttribute('list')).toMatchObject({kind:'decimal',start:1});}finally{doc.destroy();}
});

it('单元格格式跟随文字选区，独立于并发输入并可往返 DOCX',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');
 const {setWordFont,setWordLink,setWordAlignment,setWordHeading,setWordList,setWordFirstLineIndent,captureWordSelection,toggleWordMark}=await import('../packages/frontend/core/src/office-word-format');
 const {exportDocxWord}=await import('../packages/frontend/apps/web/server/office-word-export');const {encodeDocxWord,validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 const a=new Y.Doc(),b=new Y.Doc();
 try{
  const table=createWordTable(2,2);a.getArray('paragraphs').push([table]);const text=wordTableRows(table).get(0).get(0);text.insert(0,'甲乙😀');const selection=captureWordSelection(text,0,2)!;Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
  const remote=wordTableRows(b.getArray<Y.Map<unknown>>('paragraphs').get(0)).get(0).get(0);remote.insert(0,'远端');Y.applyUpdate(a,Y.encodeStateAsUpdate(b));
  setWordFont(text,selection,'fontSize',13.5);setWordFont(text,selection,'color','#123ABC');setWordLink(text,selection,'https://example.com/');toggleWordMark(text,selection,'bold');setWordAlignment(text,'center');setWordHeading(text,2);setWordList(text,'decimal');setWordFirstLineIndent(text,-360);
  Y.applyUpdate(b,Y.encodeStateAsUpdate(a));expect(text.toString()).toBe('远端甲乙😀');expect(remote.toDelta()).toEqual(text.toDelta());expect(text.toDelta()[1]).toMatchObject({insert:'甲乙',attributes:{bold:true,fontSize:13.5,color:'#123ABC',href:'https://example.com/'}});expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(a))).not.toThrow();
  const imported=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(a)));try{const restored=wordTableRows(imported.doc.getArray<Y.Map<unknown>>('paragraphs').get(0)).get(0).get(0);expect(restored.toDelta()).toEqual(text.toDelta());expect(restored.getAttributes()).toMatchObject({heading:2,align:'center',firstLineIndent:-360,list:{kind:'decimal',start:1}});}finally{imported.doc.destroy();}
 }finally{a.destroy();b.destroy();}
});
it('删除单元格所在行或表格后旧格式操作全部拒绝',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {deleteWordTableRow}=await import('../packages/frontend/core/src/office-word-table');
 const {setWordFont,setWordLink,setWordAlignment,setWordHeading,setWordList,setWordLeftIndent,setWordRightIndent,setWordFirstLineIndent,captureWordSelection,toggleWordMark}=await import('../packages/frontend/core/src/office-word-format');
 for(const removeTable of [false,true]){const doc=new Y.Doc();try{const table=createWordTable(2,2);doc.getArray('paragraphs').push([table]);const rows=wordTableRows(table),row=rows.get(0),text=row.get(0);text.insert(0,'选中文字');const selection=captureWordSelection(text,0,2)!;
  if(removeTable)doc.getArray('paragraphs').delete(0,1);else deleteWordTableRow(table,row);
  for(const action of [()=>setWordFont(text,selection,'fontSize',14),()=>setWordLink(text,selection,'https://example.com/'),()=>toggleWordMark(text,selection,'bold'),()=>setWordAlignment(text,'center'),()=>setWordHeading(text,1),()=>setWordList(text,'bullet'),()=>setWordLeftIndent(text,360),()=>setWordRightIndent(text,360),()=>setWordFirstLineIndent(text,-360)])expect(action).toThrow('删除');
 }finally{doc.destroy();}}
});

it('自定义 Word 表格尺寸和指定行前后插入保留原行文字及格式',async()=>{
 const {appendWordTable,insertWordTableRow}=await import('../packages/frontend/core/src/office-word-table');const {wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {exportDocxWord}=await import('../packages/frontend/apps/web/server/office-word-export');const {encodeDocxWord}=await import('../packages/frontend/apps/web/server/office-word');const doc=new Y.Doc();
 try{const table=appendWordTable(doc,3,2),rows=wordTableRows(table),anchor=rows.get(1);anchor.get(0).insert(0,'保留原行');anchor.get(0).format(0,2,{bold:true});const delta=anchor.get(0).toDelta();const before=insertWordTableRow(table,anchor,'before'),after=insertWordTableRow(table,anchor,'after');expect(rows.toArray()).toEqual([rows.get(0),before,anchor,after]);expect(anchor.get(0).toDelta()).toEqual(delta);expect(before.length).toBe(3);expect(after.get(2).toString()).toBe('');const imported=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(doc)));try{const result=wordTableRows(imported.doc.getArray<Y.Map<unknown>>('paragraphs').get(0));expect(result.length).toBe(4);expect(result.get(2).get(0).toDelta()).toEqual(delta);}finally{imported.doc.destroy();}}finally{doc.destroy();}
});
it('Word 并发指定行插入按对象定位收敛，已删除目标或表格拒绝插入',async()=>{
 const {appendWordTable,insertWordTableRow,deleteWordTableRow}=await import('../packages/frontend/core/src/office-word-table');const {wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');const a=new Y.Doc(),b=new Y.Doc();
 try{const ta=appendWordTable(a,2,2),ra=wordTableRows(ta),anchor=ra.get(1);anchor.get(0).insert(0,'原行');Y.applyUpdate(b,Y.encodeStateAsUpdate(a));const tb=b.getArray<Y.Map<unknown>>('paragraphs').get(0),rb=wordTableRows(tb),remoteAnchor=rb.get(1);insertWordTableRow(ta,anchor,'before').get(0).insert(0,'甲新行');insertWordTableRow(tb,remoteAnchor,'before').get(0).insert(0,'乙新行');remoteAnchor.get(0).insert(0,'远端');const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);expect(ra.toJSON()).toEqual(rb.toJSON());expect(ra.length).toBe(4);expect(ra.get(3)).toBe(anchor);expect(anchor.get(0).toString()).toBe('远端原行');expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(a))).not.toThrow();deleteWordTableRow(ta,anchor);expect(()=>insertWordTableRow(ta,anchor,'after')).toThrow('目标行已删除');a.getArray('paragraphs').delete(0,1);expect(()=>insertWordTableRow(ta,ra.get(0),'before')).toThrow('表格已删除');}finally{a.destroy();b.destroy();}
});
it('Word 新建与插入行在尺寸、行数、表格数及总单元格配额失败时不修改文档',async()=>{
 const {appendWordTable,appendWordTableRow,insertWordTableRow}=await import('../packages/frontend/core/src/office-word-table');const {wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const doc=new Y.Doc(),quota=new Y.Doc();
 try{for(const [columns,rows]of [[0,2],[51,2],[2,0],[2,201],[2,1.5],[NaN,2]])expect(()=>appendWordTable(doc,columns,rows)).toThrow();expect(doc.getArray('paragraphs').length).toBe(0);const table=appendWordTable(doc,1,200),anchor=wordTableRows(table).get(0);expect(()=>insertWordTableRow(table,anchor,'before')).toThrow('200 行');expect(wordTableRows(table).length).toBe(200);for(let i=1;i<100;i++)appendWordTable(doc,1,1);expect(()=>appendWordTable(doc,1,1)).toThrow('100 个表格');expect(doc.getArray('paragraphs').length).toBe(100);
  appendWordTable(quota,50,200);appendWordTable(quota,50,199);const last=appendWordTable(quota,50,1);expect(()=>appendWordTable(quota,1,1)).toThrow('2 万');expect(()=>appendWordTableRow(last)).toThrow('2 万');expect(wordTableRows(last).length).toBe(1);expect(quota.getArray('paragraphs').length).toBe(3);
 }finally{doc.destroy();quota.destroy();}
});

it('单元格草稿按资源独立持久化，重启后保留原操作编号、类型、基线和代际',async()=>{
 const {SpreadsheetDraftStore}=await import('../packages/frontend/core/src/spreadsheet-draft-store');
 const first=new Y.Doc(),reopened=new Y.Doc();try{
  const preferences=first.getMap<unknown>('uiPreferences'),a=new SpreadsheetDraftStore(preferences,'doc:a'),b=new SpreadsheetDraftStore(preferences,'doc:b');
  const draft={row:1,column:2,value:'001',cellType:'text' as const,base:'old-revision',generation:'epoch-original',operation:{operationId:'original-operation',type:'setCell' as const,base:'old-revision',officeGeneration:'epoch-original',value:'001',rowKey:'stable-row',columnKey:'stable-column'},error:'响应丢失',conflict:false};
  a.set('stable-cell',draft);b.set('stable-cell',{...draft,value:'其他文件'});draft.operation.operationId='accidental-mutation';
  expect(a.get('stable-cell')?.operation?.operationId).toBe('original-operation');
  Y.applyUpdate(reopened,Y.encodeStateAsUpdate(first));const saved=new SpreadsheetDraftStore(reopened.getMap('uiPreferences'),'doc:a');expect(saved.size).toBe(1);expect(saved.get('stable-cell')).toMatchObject({value:'001',generation:'epoch-original',base:'old-revision',operation:{operationId:'original-operation'},error:'响应丢失'});
  expect([...saved.keys()]).toEqual(['stable-cell']);expect(saved.has('stable-cell')).toBe(true);const seen:string[]=[];saved.forEach(value=>seen.push(value.value));expect(seen).toEqual(['001']);expect(saved.delete('missing')).toBe(false);saved.clear();expect(saved.size).toBe(0);expect(new SpreadsheetDraftStore(reopened.getMap('uiPreferences'),'doc:b').size).toBe(1);
 }finally{first.destroy();reopened.destroy();}
});

it('各格式协作位置：按类型过滤成员、忽略自身与伪造字段，最多 32 个',async()=>{
 const {officePeers}=await import('../packages/common/nbstore/src/office-presence');
 const states=new Map<number,Record<string,unknown>>([[1,{user:{name:'自己',color:'#000000'},focus:{kind:'pdf',page:1}}],[2,{user:{name:'乙',color:'#b54708'},focus:{kind:'cell',sheet:'first',row:'1-1',column:'c',label:'A1'}}],[3,{user:{name:'丙',color:'javascript:'},focus:{kind:'pdf',page:2}}],[4,{focus:{kind:'pdf',page:3}}]]);
 for(let client=10;client<60;client++)states.set(client,{user:{name:'成员'+client}});
 const peers=officePeers(states,1,'cell');expect(peers).toHaveLength(32);expect(peers[0]).toMatchObject({client:2,name:'乙',color:'#b54708',focus:{kind:'cell',label:'A1'}});expect(peers[1]).toMatchObject({client:3,color:'#3159cb'});expect(peers[1].focus).toBeUndefined();expect(peers.some(peer=>peer.client===1||peer.client===4)).toBe(false);
});

it('表格离线操作队列按创建顺序重放、断网时保留、冲突后停止后续操作并保存原编号',async()=>{
 const {SpreadsheetOperationQueue,isNetworkFailure}=await import('../packages/frontend/core/src/spreadsheet-operation-queue');const {CloudError}=await import('../packages/common/nbstore/src/cloud');
 const doc=new Y.Doc();try{const preferences=doc.getMap('ui'),queue=new SpreadsheetOperationQueue(preferences,'sheet'),other=new SpreadsheetOperationQueue(preferences,'other');
  const op=(type:string,id:string)=>({type,operationId:id.padEnd(16,'0'),base:''}) as never;const now=vi.spyOn(Date,'now');try{now.mockReturnValue(2);queue.add(op('setFormat','b'));now.mockReturnValue(1);queue.add(op('insertRow','a'));now.mockReturnValue(3);queue.add(op('mergeCells','c'));other.add(op('setFormat','z'));}finally{now.mockRestore();}
  expect(queue.entries().map(([,entry])=>entry.operation.type)).toEqual(['insertRow','setFormat','mergeCells']);expect(queue.hasStructural()).toBe(true);expect(other.size).toBe(1);
  const sent:string[]=[];expect(await queue.drain(async operation=>{if(operation.type==='setFormat')throw new TypeError('Failed to fetch');sent.push(operation.type);})).toBe(false);expect(sent).toEqual(['insertRow']);expect(queue.size).toBe(2);expect(queue.entries()[0][1].error).toBeUndefined();
  expect(await queue.drain(async operation=>{if(operation.type==='setFormat')throw new CloudError('目标单元格所属行列已删除',409);sent.push(operation.type);})).toBe(false);expect(queue.entries()[0][1]).toMatchObject({conflict:true,error:'目标单元格所属行列已删除'});expect(sent).toEqual(['insertRow']);
  expect(await queue.drain(async()=>{sent.push('never');})).toBe(false);expect(sent).toEqual(['insertRow']);queue.update(queue.entries()[0][0],{error:undefined,conflict:undefined});
  expect(await queue.drain(async operation=>{sent.push(operation.operationId.slice(0,1));})).toBe(true);expect(sent).toEqual(['insertRow','b','c']);expect(queue.size).toBe(0);expect(other.size).toBe(1);
  expect(isNetworkFailure(new TypeError('Failed to fetch'))).toBe(true);expect(isNetworkFailure(new TypeError('Load failed'))).toBe(true);expect(isNetworkFailure(new TypeError('x is not a function'))).toBe(false);expect(isNetworkFailure(new CloudError('网络',500))).toBe(false);
 }finally{doc.destroy();}
});

it('离线结构操作仅在其间只有本人重放的操作时重新定基，其他成员改动后仍返回冲突',async()=>{
 const {SpreadsheetOperationQueue}=await import('../packages/frontend/core/src/spreadsheet-operation-queue');const doc=new Y.Doc();try{
  const queue=new SpreadsheetOperationQueue(doc.getMap('ui'),'sheet'),now=vi.spyOn(Date,'now');try{now.mockReturnValue(1);queue.add({type:'setFormat',operationId:'a'.repeat(16),base:''} as never);now.mockReturnValue(2);queue.add({type:'insertRow',operationId:'b'.repeat(16),base:'S0',row:0} as never);}finally{now.mockRestore();}
  const bases:string[]=[];await queue.drain(async operation=>{bases.push(operation.base);return operation.type==='setFormat'?{before:'S0',after:'S1'}:{before:'S1',after:'S2'};});expect(bases).toEqual(['','S1']);
  const later=new SpreadsheetOperationQueue(doc.getMap('ui'),'other');later.add({type:'setFormat',operationId:'c'.repeat(16),base:''} as never);later.add({type:'insertRow',operationId:'d'.repeat(16),base:'S0',row:0} as never);
  const seen:string[]=[];await later.drain(async operation=>{seen.push(operation.base);return {before:'REMOTE',after:'S9'};});expect(seen).toEqual(['','S0']);
 }finally{doc.destroy();}
});
