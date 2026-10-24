import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {validateWordUpdate,encodeDocxWord} from '../packages/frontend/apps/web/server/office-word';
import {exportDocxWord} from '../packages/frontend/apps/web/server/office-word-export';
import {readOfficePackage,readOfficeXml} from '../packages/frontend/apps/web/server/office-package';
import {createWordToc,createWordNote,wordNoteRefs,normalizeWordNoteRefs,wordWatermarkSize,validWordWatermark,wordRoman,WORD_NOTE_CHAR,isWordToc} from '../packages/common/nbstore/src/office-word-reference';
import {wordBlockTexts,createWordTable,wordTableRows} from '../packages/common/nbstore/src/office-word-table';
import {DEFAULT_WORD_PAGE} from '../packages/common/nbstore/src/office-word-page';
import {zipSync,strToU8} from 'fflate';

const make=(...values:string[])=>{const doc=new Y.Doc(),texts=values.map(value=>new Y.Text(value));doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs').push(texts);return {doc,texts,blocks:doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs')};};
const update=(doc:Y.Doc)=>Y.encodeStateAsUpdate(doc);

it('水印校验与自动字号；目录块与注释的模型校验、编号（脚注阿拉伯数字、尾注罗马数字）与格式继承清理',()=>{
 expect(validWordWatermark({text:'机密',color:'#C0C0C0',layout:'diagonal'})).toBe(true);
 for(const bad of [{text:'',color:'#C0C0C0',layout:'diagonal'},{text:' 机密',color:'#C0C0C0',layout:'diagonal'},{text:'机密',color:'#c0c0c0',layout:'diagonal'},{text:'机密',color:'#C0C0C0',layout:'tilted'},{text:'机密',color:'#C0C0C0',layout:'diagonal',size:10},{text:'x'.repeat(51),color:'#C0C0C0',layout:'diagonal'}])expect(validWordWatermark(bad)).toBe(false);
 expect(wordWatermarkSize({text:'机密',color:'#C0C0C0',layout:'diagonal'},DEFAULT_WORD_PAGE)).toBe(144);expect(wordWatermarkSize({text:'CONFIDENTIAL DOCUMENT',color:'#C0C0C0',layout:'horizontal'},DEFAULT_WORD_PAGE)).toBeLessThan(60);expect(wordRoman(14)).toBe('xiv');
 const {doc,texts:[a,b],blocks}=make('正文一','正文二');try{
  blocks.insert(0,[createWordToc(2)]);createWordNote(doc,'n1','footnote','第一条脚注');createWordNote(doc,'n2','endnote','尾注');createWordNote(doc,'n3','footnote','第二条');createWordNote(doc,'orphan','footnote','无引用');
  a.insert(2,WORD_NOTE_CHAR,{note:'n1'});b.insert(1,WORD_NOTE_CHAR,{note:'n2',bold:true});b.insert(b.length,WORD_NOTE_CHAR,{note:'n3'});
  doc.getMap('pageSettings').set('watermark',{text:'草稿',color:'#FF0000',layout:'horizontal'});
  expect(()=>validateWordUpdate(null,update(doc))).not.toThrow();expect(wordBlockTexts(blocks.toArray())).toEqual([a,b]);
  expect(wordNoteRefs(doc,[a,b]).map(ref=>[ref.id,ref.label])).toEqual([['n1','1'],['n2','i'],['n3','2']]);
  // Typing right after a reference copies its `note` onto the new text; normalising removes it.
  a.insert(3,'续写');expect(a.toString()).toBe('正文'+WORD_NOTE_CHAR+'续写一');expect((a.toDelta() as {insert:string;attributes?:{note?:string}}[]).find(part=>part.insert.includes('续写'))?.attributes?.note).toBe('n1');expect(normalizeWordNoteRefs(a)).toBe(true);expect((a.toDelta() as {insert:string;attributes?:{note?:string}}[]).filter(part=>part.attributes?.note).map(part=>part.insert)).toEqual([WORD_NOTE_CHAR]);
  for(const [label,change]of [['目录层级',(d:Y.Doc)=>(d.getArray('paragraphs').get(0) as Y.Map<unknown>).set('levels',4)],['注释类型',(d:Y.Doc)=>(d.getMap('wordNotes').get('n1') as Y.Map<unknown>).set('kind','sidenote')],['注释格式',(d:Y.Doc)=>((d.getMap('wordNotes').get('n1') as Y.Map<unknown>).get('text') as Y.Text).format(0,1,{bold:true})],['注释编号',(d:Y.Doc)=>(d.getArray('paragraphs').get(1) as Y.Text).format(2,1,{note:'bad id!'})],['水印',(d:Y.Doc)=>d.getMap('pageSettings').set('watermark',{text:'x',color:'red',layout:'diagonal'})]] as const){
   const copy=new Y.Doc();Y.applyUpdate(copy,update(doc));change(copy);expect(()=>validateWordUpdate(null,update(copy)),label).toThrow();copy.destroy();
  }
 }finally{doc.destroy();}
});

it('DOCX 往返：文字水印写入页眉 VML，目录为 TOC 域并链接标题书签，脚注尾注写入独立部件；导入还原，Word 自身目录不当作正文',()=>{
 const {doc,texts:[title,body,sub],blocks}=make('第一章 概述','正文段落','1.1 背景');title.setAttribute('heading',1);sub.setAttribute('heading',2);try{
  blocks.insert(0,[createWordToc(3)]);createWordNote(doc,'f1','footnote','数据来自 2026 年统计 & <年报>');createWordNote(doc,'e1','endnote','参考文献一\n第二行');
  body.insert(4,WORD_NOTE_CHAR,{note:'f1'});body.insert(body.length,WORD_NOTE_CHAR,{note:'e1'});
  const table=createWordTable(1,1);blocks.push([table]);createWordNote(doc,'f2','footnote','表格中的脚注');wordTableRows(table).get(0).get(0).insert(0,'单元格'+WORD_NOTE_CHAR,{});wordTableRows(table).get(0).get(0).format(3,1,{note:'f2'});
  doc.getMap('pageSettings').set('watermark',{text:'机密',color:'#FF0000',layout:'diagonal'});
  const bytes=exportDocxWord(update(doc)),files=readOfficePackage(bytes),xml=(name:string)=>readOfficeXml(files[name]);
  const documentXml=xml('word/document.xml');
  expect(documentXml).toContain('<w:docPartGallery w:val="Table of Contents"/>');expect(documentXml).toContain('TOC \\o "1-3" \\h \\z \\u');expect(documentXml).toMatch(/<w:hyperlink w:anchor="_Toc100000000" w:history="1"><w:r><w:t xml:space="preserve">第一章 概述<\/w:t>/);expect(documentXml).toContain('<w:bookmarkStart w:id="900001" w:name="_Toc100000001"/>');
  expect(documentXml).toContain('<w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/>');expect(documentXml).toContain('<w:endnoteReference w:id="1"/>');expect(documentXml).toContain('<w:footnoteReference w:id="2"/>');expect(documentXml).not.toContain(WORD_NOTE_CHAR);
  expect(xml('word/footnotes.xml')).toContain('<w:footnote w:type="separator" w:id="-1">');expect(xml('word/footnotes.xml')).toContain('数据来自 2026 年统计 &amp; &lt;年报&gt;');expect(xml('word/endnotes.xml')).toContain('第二行');
  expect(xml('word/settings.xml')).toContain('<w:updateFields w:val="true"/>');expect(xml('word/_rels/document.xml.rels')).toContain('relationships/footnotes');expect(xml('[Content_Types].xml')).toContain('footnotes+xml');
  const header=xml('word/header1.xml');expect(header).toContain('id="PowerPlusWaterMarkObject1"');expect(header).toContain('string="机密"');expect(header).toContain('rotation:315');expect(header).toContain('fillcolor="#ff0000"');
  const back=encodeDocxWord(bytes);try{
   const restored=back.doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs').toArray();expect(isWordToc(restored[0])).toBe(true);expect((restored[0] as Y.Map<unknown>).get('levels')).toBe(3);expect((restored[1] as Y.Text).toString()).toBe('第一章 概述');
   expect(back.doc.getMap('pageSettings').get('watermark')).toEqual({text:'机密',color:'#FF0000',layout:'diagonal'});
   const refs=wordNoteRefs(back.doc,wordBlockTexts(restored));expect(refs.map(ref=>[ref.kind,ref.label,String((back.doc.getMap('wordNotes').get(ref.id) as Y.Map<unknown>).get('text'))])).toEqual([['footnote','1','数据来自 2026 年统计 & <年报>'],['endnote','i','参考文献一\n第二行'],['footnote','2','表格中的脚注']]);
   expect((restored[2] as Y.Text).toString()).toBe('正文段落'+WORD_NOTE_CHAR+WORD_NOTE_CHAR);
  }finally{back.doc.destroy();}
  // A Word TOC field without a content control, spanning paragraphs, is also replaced by one block.
  const word=zipSync({'word/document.xml':strToU8('<w:document xmlns:w="w"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-2" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>旧目录项</w:t></w:r></w:p><w:p><w:r><w:t>第二项</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:r><w:t>正文</w:t></w:r></w:p></w:body></w:document>')});
  const parsed=encodeDocxWord(word);try{const list=parsed.doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs').toArray();expect(list).toHaveLength(2);expect((list[0] as Y.Map<unknown>).get('levels')).toBe(2);expect((list[1] as Y.Text).toString()).toBe('正文');}finally{parsed.doc.destroy();}
 }finally{doc.destroy();}
});
