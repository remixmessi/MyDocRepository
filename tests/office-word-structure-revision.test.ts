import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {strToU8,zipSync} from 'fflate';
import {formatWordTracked,setWordParagraphAttributeTracked,markWordParagraphInserted,deleteWordParagraphBreak,markWordImageInserted,deleteWordImageTracked,resolveWordRevisions,wordRevisions,setWordTrackingAuthor} from '../packages/common/nbstore/src/office-word-revision';
import {createWordImage} from '../packages/common/nbstore/src/office-image';
import {validateWordUpdate,encodeDocxWord} from '../packages/frontend/apps/web/server/office-word';
import {exportDocxWord} from '../packages/frontend/apps/web/server/office-word-export';
import {packOfficeMedia,mediaDigest} from '../packages/frontend/apps/web/server/office-media';
import {readOfficePackage,readOfficeXml} from '../packages/frontend/apps/web/server/office-package';
import {splitWordParagraph,mergeWordParagraph} from '../packages/frontend/core/src/office-word-flow';
import {toggleWordMark,captureWordSelection,setWordAlignment} from '../packages/frontend/core/src/office-word-format';

const alice={id:'alice',name:'王小明'},bob={id:'bob',name:'Lisa'};
const make=(...values:string[])=>{const doc=new Y.Doc(),texts=values.map(value=>new Y.Text(value));doc.getArray<Y.Text>('paragraphs').push(texts);return {doc,texts,blocks:doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs')};};
const runs=(text:Y.Text)=>(text.toDelta() as {insert:string;attributes?:Record<string,unknown>}[]).map(part=>({text:part.insert,...part.attributes}) as {text:string}&Record<string,unknown>);
const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNQTX79HwAElwJzVt2vfAAAAABJRU5ErkJggg==','base64');
const logo={id:'pic',name:'pic.png',mime:'image/png',data:png.toString('base64'),sha256:mediaDigest(png)};

it('格式修订：记录更改前的值，接受保留新格式，拒绝恢复原格式；改回原值撤销记录；自己插入的文字不另记格式修订',()=>{
 const {doc,texts:[text]}=make('甲乙丙丁');text.format(2,2,{italic:true});const t0=1_000_000;
 doc.transact(()=>formatWordTracked(text,0,4,{bold:true},alice,t0));
 expect(runs(text).map(({text,bold,italic,fmt})=>[text,bold,italic,(fmt as {old:unknown}|undefined)?.old])).toEqual([['甲乙',true,undefined,{bold:null}],['丙丁',true,true,{bold:null}]]);
 // Runs with the same revision and previous values form one entry.
 const [revision]=wordRevisions(doc);expect(revision).toMatchObject({kind:'fmt',author:'alice',start:0,end:4,old:{bold:null}});
 expect(new Set(wordRevisions(doc).map(item=>item.id)).size).toBe(1);
 // Same author changes more keys: the original values stay recorded once.
 doc.transact(()=>formatWordTracked(text,0,2,{fontSize:14},alice,t0+10));expect((runs(text)[0].fmt as {old:unknown}).old).toEqual({bold:null,fontSize:null});
 validateWordUpdate(null,Y.encodeStateAsUpdate(doc));
 // Reject restores the old formatting and removes the record.
 const copy=new Y.Doc();Y.applyUpdate(copy,Y.encodeStateAsUpdate(doc));expect(resolveWordRevisions(copy,false,null)).toBe(1);
 expect(runs(copy.getArray<Y.Text>('paragraphs').get(0))).toEqual([{text:'甲乙'},{text:'丙丁',italic:true}]);
 // Accept keeps the new formatting.
 resolveWordRevisions(doc,true,null);expect(runs(text)).toEqual([{text:'甲乙',bold:true,fontSize:14},{text:'丙丁',bold:true,italic:true}]);
 // Setting a value back to the original drops the record.
 const other=make('文字').texts[0];other.doc!.transact(()=>{formatWordTracked(other,0,2,{bold:true},alice,t0);formatWordTracked(other,0,2,{bold:null},alice,t0+5);});expect(runs(other)).toEqual([{text:'文字'}]);
 // Own pending insertions are formatted without a separate record.
 const own=make('').texts[0];own.insert(0,'新',{ins:{id:'r1',author:'alice',name:'王小明',date:t0}});own.doc!.transact(()=>formatWordTracked(own,0,1,{bold:true},alice,t0));expect(runs(own)[0].fmt).toBeUndefined();
});

it('段落格式修订：pfmt 记录原对齐/标题，拒绝恢复；编辑器格式函数在修订模式下自动记录',()=>{
 const {doc,texts:[text]}=make('标题');text.setAttribute('align','left');
 doc.transact(()=>{setWordParagraphAttributeTracked(text,'heading',1,alice);setWordParagraphAttributeTracked(text,'align','center',alice);});
 expect(text.getAttribute('pfmt')).toMatchObject({author:'alice',old:{heading:null,align:'left'}});validateWordUpdate(null,Y.encodeStateAsUpdate(doc));
 resolveWordRevisions(doc,false,null);expect(text.getAttributes()).toEqual({align:'left'});
 // Through the editor's helpers: the tracking author registered for the document is used.
 setWordTrackingAuthor(doc,bob);try{setWordAlignment(text,'right');toggleWordMark(text,captureWordSelection(text,0,2)!,'bold');}finally{setWordTrackingAuthor(doc,undefined);}
 expect(wordRevisions(doc).map(item=>[item.kind,item.author])).toEqual([['pfmt','bob'],['fmt','bob']]);
});

it('分段与合段修订：分段标记 insBreak，合段标记 delBreak 且段落保持分开；拒绝分段与接受合段时真正合并；撤销自己的分段直接合并',()=>{
 const {doc,texts:[first],blocks}=make('第一段第二段');setWordTrackingAuthor(doc,alice);
 try{
  const second=splitWordParagraph(first,3);expect(blocks.length).toBe(2);expect(second.getAttribute('insBreak')).toMatchObject({author:'alice'});
  // Alice's own pending split: Backspace at the start of the new paragraph merges it back.
  const merged=mergeWordParagraph(second,'backward');expect(merged).toEqual({text:first,offset:3});expect(blocks.length).toBe(1);expect(first.toString()).toBe('第一段第二段');
 }finally{setWordTrackingAuthor(doc,undefined);}
 // Bob splits; Alice rejects → merged again.
 setWordTrackingAuthor(doc,bob);const second=splitWordParagraph(first,3);setWordTrackingAuthor(doc,undefined);
 expect(wordRevisions(doc).map(item=>item.kind)).toEqual(['insBreak']);validateWordUpdate(null,Y.encodeStateAsUpdate(doc));
 resolveWordRevisions(doc,false,null);expect(blocks.length).toBe(1);expect(first.toString()).toBe('第一段第二段');expect(second._item?.deleted).toBe(true);
 // Joining two existing paragraphs marks the break deleted; both stay until accepted.
 const two=make('上','下');setWordTrackingAuthor(two.doc,alice);
 try{const caret=mergeWordParagraph(two.texts[1],'backward');expect(caret).toEqual({text:two.texts[0],offset:1});expect(two.blocks.length).toBe(2);expect(two.texts[1].getAttribute('delBreak')).toMatchObject({author:'alice'});
  // Pressing Backspace again does not stack another revision.
  mergeWordParagraph(two.texts[1],'backward');expect(wordRevisions(two.doc)).toHaveLength(1);}finally{setWordTrackingAuthor(two.doc,undefined);}
 const reject=new Y.Doc();Y.applyUpdate(reject,Y.encodeStateAsUpdate(two.doc));resolveWordRevisions(reject,false,null);expect(reject.getArray('paragraphs').length).toBe(2);
 resolveWordRevisions(two.doc,true,null);expect(two.blocks.length).toBe(1);expect(two.texts[0].toString()).toBe('上下');
 // Paragraph-mark revisions are body-only: cells and malformed marks are refused.
 const forged=make('a','b');forged.texts[1].setAttribute('delBreak',{id:'x',author:'a'});expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(forged.doc))).toThrow('段落属性');
});

it('分段时选中的文字在修订模式下标为删除并留在前一段',()=>{
 const {doc,texts:[text],blocks}=make('甲乙丙丁');setWordTrackingAuthor(doc,alice);
 try{const next=splitWordParagraph(text,1,3);expect(text.toString()).toBe('甲乙丙');expect(runs(text).map(run=>[run.text,!!run.del])).toEqual([['甲',false],['乙丙',true]]);expect(next.toString()).toBe('丁');expect(blocks.length).toBe(2);}
 finally{setWordTrackingAuthor(doc,undefined);}
});

it('图片修订：插入带 ins、删除他人图片标 del、删除自己待审图片直接移除；接受/拒绝；只有正文图片可带修订',()=>{
 const {doc,blocks}=make('正文');const mine=createWordImage('pic',10,10,'新图'),theirs=createWordImage('pic',10,10,'旧图');markWordImageInserted(mine,alice);blocks.push([mine,theirs]);
 expect(deleteWordImageTracked(mine,alice)).toBe(true);expect(deleteWordImageTracked(theirs,alice)).toBe(false);expect(theirs.get('del')).toMatchObject({author:'alice'});
 validateWordUpdate(null,Y.encodeStateAsUpdate(doc));expect(wordRevisions(doc).map(item=>item.kind)).toEqual(['insImage','delImage']);
 const reject=new Y.Doc();Y.applyUpdate(reject,Y.encodeStateAsUpdate(doc));resolveWordRevisions(reject,false,null);expect(reject.getArray('paragraphs').length).toBe(2);
 expect((reject.getArray('paragraphs').get(1) as Y.Map<unknown>).has('del')).toBe(false);
 resolveWordRevisions(doc,true,null);expect(blocks.length).toBe(2);expect((blocks.get(1) as Y.Map<unknown>).get('alt')).toBe('新图');expect((blocks.get(1) as Y.Map<unknown>).has('ins')).toBe(false);
 // Header pictures cannot carry revisions.
 const header=make('x'),pictures=new Y.Map<Y.Map<unknown>>(),image=createWordImage('pic',10,10,'标志');markWordImageInserted(image,alice);pictures.set('main.header.left',image);header.doc.getMap('pageSettings').set('headerImages',pictures);
 expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(header.doc))).toThrow('修订');
});

it('团队文档：新的格式、分段与图片修订必须属于当前成员',()=>{
 const {doc}=make('文字'),before=Y.encodeStateAsUpdate(doc);doc.destroy();
 const cases:((t:Y.Text)=>void)[]=[t=>formatWordTracked(t,0,1,{bold:true},bob),t=>markWordParagraphInserted(t,bob),t=>setWordParagraphAttributeTracked(t,'align','center',bob)];
 for(const edit of cases){const copy=new Y.Doc();Y.applyUpdate(copy,before);copy.transact(()=>edit(copy.getArray<Y.Text>('paragraphs').get(0)));const update=Y.encodeStateAsUpdate(copy);
  expect(()=>validateWordUpdate(before,update,{actor:'alice'})).toThrow('作者');expect(()=>validateWordUpdate(before,update,{actor:'bob'})).not.toThrow();copy.destroy();}
});

it('DOCX 往返：w:rPrChange / w:pPrChange 写入原格式，段落标记修订写在前一段 pPr/rPr，图片包在 w:ins/w:del 中；重新导入还原，LibreOffice 格式亦可读',()=>{
 const {doc,texts:[a,b],blocks}=make('加粗文字','第二段');doc.getMap('pageSettings').set('trackChanges',true);
 doc.transact(()=>{formatWordTracked(a,0,2,{bold:true,fontSize:16},alice,1_700_000_000_000);setWordParagraphAttributeTracked(a,'align','center',alice,1_700_000_000_000);});
 b.setAttribute('delBreak',{id:'brk',author:'bob',name:'Lisa',date:1_700_000_000_000});
 const image=createWordImage('pic',20,20,'插图');markWordImageInserted(image,bob,1_700_000_000_000);blocks.push([image]);
 const bytes=exportDocxWord(packOfficeMedia(Y.encodeStateAsUpdate(doc),[logo])),xml=readOfficeXml(readOfficePackage(bytes)['word/document.xml']);
 expect(xml).toMatch(/<w:rPr><w:b w:val="1"\/><w:sz w:val="32"\/><w:rPrChange w:id="\d+" w:author="王小明" w:date="2023-11-14T22:13:20Z"><w:rPr><\/w:rPr><\/w:rPrChange><\/w:rPr>/);
 expect(xml).toMatch(/<w:jc w:val="center"\/><w:rPr><w:del w:id="\d+" w:author="Lisa" w:date="[^"]+"\/><\/w:rPr><w:pPrChange w:id="\d+" w:author="王小明" w:date="[^"]+"><w:pPr><w:pStyle w:val="Normal"\/><\/w:pPr><\/w:pPrChange><\/w:pPr>/);
 expect(xml).toMatch(/<w:p><w:ins w:id="\d+" w:author="Lisa" w:date="[^"]+"><w:r><w:drawing>/);
 const back=encodeDocxWord(bytes);try{
  expect(wordRevisions(back.doc).map(item=>[item.kind,item.name,item.old??null])).toEqual([['pfmt','王小明',{align:null}],['fmt','王小明',{bold:null,fontSize:null}],['delBreak','Lisa',null],['insImage','Lisa',null]]);
  resolveWordRevisions(back.doc,false,null);const paragraphs=back.doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs');expect(paragraphs.length).toBe(2);expect(runs(paragraphs.get(0) as Y.Text)).toEqual([{text:'加粗文字'}]);expect((paragraphs.get(0) as Y.Text).getAttributes()).toEqual({});
 }finally{back.doc.destroy();}
 // Word's own markup: rPrChange with old properties, an inserted paragraph mark and pPrChange.
 const w='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
 const files={'word/document.xml':strToU8(`<?xml version="1.0"?><w:document ${w}><w:body><w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:ins w:id="7" w:author="Word 用户" w:date="2026-10-22T08:00:00Z"/></w:rPr><w:pPrChange w:id="8" w:author="Word 用户" w:date="2026-10-22T08:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:i/><w:rPrChange w:id="9" w:author="Word 用户" w:date="2026-10-22T08:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>上</w:t></w:r></w:p><w:p><w:r><w:t>下</w:t></w:r></w:p></w:body></w:document>`)};
 const word=encodeDocxWord(zipSync(files));try{
  expect(wordRevisions(word.doc).map(item=>[item.kind,item.old??null])).toEqual([['pfmt',{align:'left'}],['fmt',{bold:true,italic:null}],['insBreak',null]]);
  resolveWordRevisions(word.doc,false,null);expect(word.doc.getArray('paragraphs').length).toBe(1);expect(runs(word.doc.getArray<Y.Text>('paragraphs').get(0))).toEqual([{text:'上',bold:true},{text:'下'}]);
 }finally{word.doc.destroy();}
});
