import {expect,it} from 'vitest';
import * as Y from 'yjs';
import {exportDocxWord} from '../packages/frontend/apps/web/server/office-word-export';
import {encodeDocxWord} from '../packages/frontend/apps/web/server/office-word';
import {readOfficePackage,readOfficeXml} from '../packages/frontend/apps/web/server/office-package';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';

function snapshot(values:string[]){const doc=new Y.Doc();doc.getArray<Y.Text>('paragraphs').push(values.map(value=>new Y.Text(value)));return doc;}
it('生成最小 OOXML 包并保留中文、emoji、空白、制表符、换行和空段落',()=>{
 const values=['  中文😀 &<>"\'\t第二列\n第二行\r末尾  ','','最后一段'],doc=snapshot(values);
 try{
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes);
  expect(Object.keys(files).sort()).toEqual(['[Content_Types].xml','_rels/.rels','word/_rels/document.xml.rels','word/document.xml','word/styles.xml']);
  expect(readOfficeXml(files['[Content_Types].xml'])).toContain('wordprocessingml.document.main+xml');
  expect(readOfficeXml(files['_rels/.rels'])).toContain('Target="word/document.xml"');
  const imported=encodeDocxWord(bytes);try{expect(imported.doc.getArray<Y.Text>('paragraphs').toArray().map(p=>p.toString())).toEqual(values);}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('拒绝非法 XML 字符、富文本和嵌入，不静默丢失',()=>{
 for(const value of ['\u0000','\uffff']){const doc=snapshot([value]);try{expect(()=>exportDocxWord(Y.encodeStateAsUpdate(doc))).toThrow('字符');}finally{doc.destroy();}}
 const doc=snapshot(['格式']);try{doc.getArray<Y.Text>('paragraphs').get(0).format(0,1,{color:'red'});expect(()=>exportDocxWord(Y.encodeStateAsUpdate(doc))).toThrow('富文本');}finally{doc.destroy();}
 const embedded=snapshot(['']);try{embedded.getArray<Y.Text>('paragraphs').get(0).insertEmbed(0,{image:'x'});expect(()=>exportDocxWord(Y.encodeStateAsUpdate(embedded))).toThrow('嵌入');}finally{embedded.destroy();}
});
it('空模型可导出，超限正文被明确拒绝',()=>{
 const empty=snapshot([]);try{const parsed=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(empty)));expect(parsed.doc.getArray('paragraphs').length).toBe(1);parsed.doc.destroy();}finally{empty.destroy();}
 const large=snapshot(Array.from({length:90},()=>'<'.repeat(100000)));try{expect(()=>exportDocxWord(Y.encodeStateAsUpdate(large))).toThrow('8 MB');}finally{large.destroy();}
});
it('DOCX 导出复用隔离进程并可重新解析',async()=>{
 const pool=new OfficeParserPool(),doc=snapshot(['子进程\t导出\n保留']);
 try{const bytes=await pool.exportDocx(Y.encodeStateAsUpdate(doc));const update=await pool.parse('docx',bytes),result=new Y.Doc();try{Y.applyUpdate(result,update);expect(result.getArray<Y.Text>('paragraphs').get(0).toString()).toBe('子进程\t导出\n保留');}finally{result.destroy();}expect(pool.activeCount).toBe(0);}finally{doc.destroy();pool.close();}
});

it('直接文字样式跨导入导出保留，关闭样式不被错误打开',async()=>{
 const {zipSync,strToU8}=await import('fflate');
 const source=zipSync({'word/document.xml':strToU8(`<w:document><w:body><w:p><w:r><w:rPr><w:b/><w:i w:val="on"/><w:u w:val="single"/></w:rPr><w:t>粗斜线😀</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:i w:val="false"/><w:u w:val="none"/></w:rPr><w:t>普通\t</w:t></w:r><w:r><w:t>未设置</w:t></w:r></w:p></w:body></w:document>`)});
 const imported=encodeDocxWord(source);
 try{
  const before=imported.doc.getArray<Y.Text>('paragraphs').get(0).toDelta();
  expect(before).toEqual([{insert:'粗斜线😀',attributes:{bold:true,italic:true,underline:true}},{insert:'普通\t',attributes:{bold:false,italic:false,underline:false}},{insert:'未设置'}]);
  const output=exportDocxWord(imported.update),reopened=encodeDocxWord(output);
  try{expect(reopened.doc.getArray<Y.Text>('paragraphs').get(0).toDelta()).toEqual(before);}finally{reopened.doc.destroy();}
 }finally{imported.doc.destroy();}
});
it('合法样式可提交，非法属性、段落属性与嵌入被拒绝',async()=>{
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const attributes of [{bold:true,italic:false,underline:true},{bold:'true'},{fontFamily:'a;b{}'},{shadow:true},{highlight:'orange'},{vertAlign:'top'}] as Record<string,unknown>[]){
  const doc=snapshot(['正文']);try{doc.getArray<Y.Text>('paragraphs').get(0).format(0,2,attributes);if(typeof attributes.bold==='boolean')expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();else expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('属性');}finally{doc.destroy();}
 }
 const doc=snapshot(['正文']);try{doc.getArray<Y.Text>('paragraphs').get(0).setAttribute('style','hidden');expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('段落属性');}finally{doc.destroy();}
});

it('一级至三级标题及空标题往返，保留文字片段属性和正文',()=>{
 const doc=snapshot(['一级😀','二级','', '正文']);
 try{
  const paragraphs=doc.getArray<Y.Text>('paragraphs');for(let i=0;i<3;i++)paragraphs.get(i).setAttribute('heading',i+1);
  paragraphs.get(0).format(0,2,{bold:false,italic:true});
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes);
  expect(readOfficeXml(files['word/_rels/document.xml.rels'])).toContain('Target="styles.xml"');
  expect(readOfficeXml(files['[Content_Types].xml'])).toContain('wordprocessingml.styles+xml');
  const reopened=encodeDocxWord(bytes);try{
   const restored=reopened.doc.getArray<Y.Text>('paragraphs').toArray();
   expect(restored.map(p=>p.getAttribute('heading')??0)).toEqual([1,2,3,0]);expect(restored.map(p=>p.toString())).toEqual(['一级😀','二级','','正文']);expect(restored[0].toDelta()).toEqual(paragraphs.get(0).toDelta());
  }finally{reopened.doc.destroy();}
 }finally{doc.destroy();}
});
it('导入显式大纲、基于自定义样式的标题及覆盖为正文',async()=>{
 const {zipSync,strToU8}=await import('fflate');
 const styles='<w:styles><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Custom"><w:basedOn w:val="Heading1"/></w:style><w:style w:type="paragraph" w:styleId="Other"><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style></w:styles>';
 const properties=['<w:pStyle w:val="Custom"/>','<w:pStyle w:val="Custom"/><w:outlineLvl w:val="9"/>','<w:pStyle w:val="Other"/>','<w:outlineLvl w:val="1"/>','<w:pStyle w:val="Heading2"/>','<w:outlineLvl w:val="8"/>'];
 const bytes=zipSync({'word/styles.xml':strToU8(styles),'word/document.xml':strToU8('<w:document><w:body>'+properties.map(value=>'<w:p><w:pPr>'+value+'</w:pPr><w:r><w:t>文字</w:t></w:r></w:p>').join('')+'</w:body></w:document>')});
 const parsed=encodeDocxWord(bytes);try{expect(parsed.doc.getArray<Y.Text>('paragraphs').toArray().map(p=>p.getAttribute('heading')??0)).toEqual([1,0,3,2,2,0]);}finally{parsed.doc.destroy();}
});
it('拒绝样式继承循环、过深、重复编号及非法段落标题属性',async()=>{
 const {wordHeadingReader}=await import('../packages/frontend/apps/web/server/office-word-heading');
 const xml=(value:string)=>new TextEncoder().encode('<w:styles>'+value+'</w:styles>');
 expect(()=>wordHeadingReader(xml('<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>'))).toThrow('循环');
 expect(()=>wordHeadingReader(xml('<w:style w:type="paragraph" w:styleId="A"/><w:style w:type="paragraph" w:styleId="A"/>'))).toThrow('重复');
 expect(()=>wordHeadingReader(xml(Array.from({length:66},(_,i)=>`<w:style w:type="paragraph" w:styleId="S${i}"><w:basedOn w:val="S${i+1}"/></w:style>`).join('')))).toThrow('64');
 expect(()=>wordHeadingReader(xml(Array.from({length:66},(_,i)=>`<w:style w:type="paragraph" w:styleId="S${i}"><w:basedOn w:val="S${i+1}"/></w:style>`).reverse().join('')))).toThrow('64');
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const heading of [-1,4,1.5,'1',true,{},null]){const doc=snapshot(['正文']);try{doc.getArray<Y.Text>('paragraphs').get(0).setAttribute('heading',heading);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('段落属性');}finally{doc.destroy();}}
});

it('单层圆点和编号导出包含关系，空项、标题和文字样式可往返',()=>{
 const doc=snapshot(['第一项','','正文','继续编号','圆点']);
 try{
  const p=doc.getArray<Y.Text>('paragraphs');for(const i of [0,1,3])p.get(i).setAttribute('list',{id:'group-a',kind:'decimal',start:3});p.get(4).setAttribute('list',{id:'group-b',kind:'bullet',start:1});p.get(0).setAttribute('heading',2);p.get(0).format(0,2,{bold:false,underline:true});
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes);
  expect(readOfficeXml(files['word/numbering.xml'])).toContain('<w:start w:val="3"/>');expect(readOfficeXml(files['word/_rels/document.xml.rels'])).toContain('Target="numbering.xml"');
  expect(readOfficeXml(files['[Content_Types].xml'])).toContain('wordprocessingml.numbering+xml');
  const imported=encodeDocxWord(bytes);try{
   const paragraphs=imported.doc.getArray<Y.Text>('paragraphs').toArray();expect(paragraphs.map(item=>item.toString())).toEqual(['第一项','','正文','继续编号','圆点']);
   const list=paragraphs[0].getAttribute('list');expect(list).toMatchObject({kind:'decimal',start:3});expect(paragraphs[1].getAttribute('list')).toEqual(list);expect(paragraphs[3].getAttribute('list')).toEqual(list);expect(paragraphs[2].getAttribute('list')).toBeUndefined();expect(paragraphs[4].getAttribute('list')).toMatchObject({kind:'bullet'});expect(paragraphs[0].getAttribute('heading')).toBe(2);expect(paragraphs[0].toDelta()).toEqual(p.get(0).toDelta());
  }finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('读取编号起始覆盖和样式继承，拒绝多层、缺失和不支持的编号格式',async()=>{
 const {wordListReader}=await import('../packages/frontend/apps/web/server/office-word-list');const bytes=(xml:string)=>new TextEncoder().encode(xml);
 const numbering=bytes('<w:numbering><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>');
 const styles=bytes('<w:styles><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Custom"><w:basedOn w:val="Base"/></w:style></w:styles>');
 const read=wordListReader(numbering,styles),p=(num:string,level=0)=>`<w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${num}"/></w:numPr></w:pPr>`;
 expect(read(p('7'))).toEqual({kind:'decimal',id:'import-7',start:5});expect(read('<w:pPr><w:pStyle w:val="Custom"/></w:pPr>')).toMatchObject({start:5});expect(read(p('0'))).toBeUndefined();expect(()=>read(p('7',1))).toThrow('多级');expect(()=>read(p('8'))).toThrow('缺失');
 expect(()=>wordListReader(bytes('<w:numbering><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"></w:abstractNum></w:numbering>'))).toThrow('重复');
 const unsupported=wordListReader(bytes(new TextDecoder().decode(numbering).replace('decimal','upperRoman')));expect(()=>unsupported(p('7'))).toThrow('格式');
});
it('拒绝非法列表属性和同组矛盾定义',async()=>{
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const list of [{id:'x',kind:'roman',start:1},{id:'x',kind:'decimal',start:0},{id:'x',kind:'decimal',start:1,hidden:true},'bullet']){const doc=snapshot(['文字']);try{doc.getArray<Y.Text>('paragraphs').get(0).setAttribute('list',list);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('段落属性');}finally{doc.destroy();}}
 const doc=snapshot(['一','二']);try{doc.getArray<Y.Text>('paragraphs').get(0).setAttribute('list',{id:'same',kind:'decimal',start:1});doc.getArray<Y.Text>('paragraphs').get(1).setAttribute('list',{id:'same',kind:'decimal',start:2});expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('不一致');}finally{doc.destroy();}
});

it('网页链接与普通文字顺序、字符格式、中文和 XML 特殊字符往返',()=>{
 const doc=snapshot(['前链接😀后链接']),href='https://example.com/?q="x"&lang=中文#片段';
 try{
  const p=doc.getArray<Y.Text>('paragraphs').get(0);p.format(1,4,{href,bold:true,underline:false});p.format(6,2,{href,italic:true});
  const files=readOfficePackage(exportDocxWord(Y.encodeStateAsUpdate(doc))),rels=readOfficeXml(files['word/_rels/document.xml.rels']);
  expect(rels.match(/TargetMode="External"/g)).toHaveLength(1);expect(rels).toContain('&quot;x&quot;&amp;lang=中文');
  const imported=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(doc)));try{expect(imported.doc.getArray<Y.Text>('paragraphs').get(0).toDelta()).toEqual(p.toDelta());}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('导入拒绝危险地址、缺失或错误链接关系及未支持的内部链接和域代码',async()=>{
 const {zipSync,strToU8}=await import('fflate');
 const source=(target:string,extra='',relationship=true)=>zipSync({'word/document.xml':strToU8(`<w:document><w:body><w:p><w:hyperlink r:id="rId1" ${extra}><w:r><w:t>链接</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`),...(relationship?{'word/_rels/document.xml.rels':strToU8(`<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${target}" TargetMode="External"/></Relationships>`)}:{})});
 for(const href of ['javascript:alert(1)','file:///etc/passwd','data:text/html,test','mailto:a@example.com','https://user:password@example.com/'])expect(()=>encodeDocxWord(source(href))).toThrow('HTTP/HTTPS');
 expect(()=>encodeDocxWord(source('https://example.com/','',false))).toThrow('缺失');expect(()=>encodeDocxWord(source('https://example.com/','w:anchor="bookmark"'))).toThrow('书签');
 expect(()=>encodeDocxWord(zipSync({'word/document.xml':strToU8('<w:document><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:document>')}))).toThrow('域代码');
});
it('同步校验拒绝不安全或超长链接属性，允许 HTTP/HTTPS',async()=>{
 const {validWordLink}=await import('../packages/common/nbstore/src/office-word-link');
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const href of ['javascript:alert(1)','https://user@example.com/','https://example.com/\n','https://example.com/'+ 'x'.repeat(2048),true]){expect(validWordLink(href)).toBe(false);const doc=snapshot(['链接']);try{doc.getArray<Y.Text>('paragraphs').get(0).format(0,2,{href});expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('富文本');}finally{doc.destroy();}}
 expect(validWordLink('https://example.com/\ud800')).toBe(false);
 expect(validWordLink('http://example.com/')).toBe(true);expect(validWordLink('https://example.com/?a=1&b=2')).toBe(true);
});

it('四种对齐及非标准左缩进往返，并保留标题列表链接和直接字符格式',()=>{
 const doc=snapshot(['左','中','右','两端']);
 try{
  const paragraphs=doc.getArray<Y.Text>('paragraphs');['left','center','right','justify'].forEach((align,i)=>{paragraphs.get(i).setAttribute('align',align);paragraphs.get(i).setAttribute('leftIndent',[0,123,720,2880][i]);});
  paragraphs.get(1).setAttribute('heading',2);paragraphs.get(1).setAttribute('list',{kind:'decimal',id:'a',start:1});paragraphs.get(1).format(0,1,{href:'https://example.com/',bold:false});
  const imported=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(doc)));try{
   const result=imported.doc.getArray<Y.Text>('paragraphs');expect(result.toArray().map(p=>p.getAttribute('align'))).toEqual(['left','center','right','justify']);expect(result.toArray().map(p=>p.getAttribute('leftIndent'))).toEqual([0,123,720,2880]);expect(result.get(1).getAttribute('heading')).toBe(2);expect(result.get(1).getAttribute('list')).toMatchObject({kind:'decimal'});expect(result.get(1).toDelta()).toEqual(paragraphs.get(1).toDelta());
  }finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('按段落样式继承对齐和缩进，直接零值覆盖继承，未知模式及超限值拒绝',async()=>{
 const {wordLayoutReader}=await import('../packages/frontend/apps/web/server/office-word-layout');const bytes=(s:string)=>new TextEncoder().encode(s);
 const read=wordLayoutReader(bytes('<w:styles><w:style w:type="paragraph" w:styleId="Base" w:default="1"><w:pPr><w:ind w:left="720"/><w:jc w:val="right"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Custom"><w:basedOn w:val="Base"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style></w:styles>'));
 expect(read('')).toEqual({align:'right',leftIndent:720});expect(read('<w:pPr><w:pStyle w:val="Custom"/></w:pPr>')).toEqual({align:'center',leftIndent:720});expect(read('<w:pPr><w:pStyle w:val="Custom"/><w:ind w:left="0"/><w:jc w:val="left"/></w:pPr>')).toEqual({align:'left',leftIndent:0});
 for(const align of ['distribute','__proto__','toString'])expect(()=>read(`<w:pPr><w:jc w:val="${align}"/></w:pPr>`)).toThrow('对齐');
 for(const indent of ['-1','2881','1.5','Infinity'])expect(()=>read(`<w:pPr><w:ind w:left="${indent}"/></w:pPr>`)).toThrow('缩进');
});
it('通用更新拒绝非法对齐和左缩进属性',async()=>{
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const [key,value]of [['align','start'],['align',true],['leftIndent','720'],['leftIndent',-1],['leftIndent',2881],['leftIndent',1.5]] as [string,unknown][]){const doc=snapshot(['文']);try{doc.getArray<Y.Text>('paragraphs').get(0).setAttribute(key,value);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('段落属性');}finally{doc.destroy();}}
});

it('右缩进和互斥的首行/悬挂缩进往返，保留零值、非标准值和字符格式',()=>{
 const doc=snapshot(['首行','悬挂','零值','非标准']);
 try{
  const p=doc.getArray<Y.Text>('paragraphs');[720,-720,0,-123].forEach((value,i)=>{p.get(i).setAttribute('firstLineIndent',value);p.get(i).setAttribute('rightIndent',[360,2880,0,123][i]);});p.get(0).format(0,2,{bold:true});
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),xml=readOfficeXml(readOfficePackage(bytes)['word/document.xml']);expect(xml).toContain('w:firstLine="720"');expect(xml).toContain('w:hanging="720"');
  const imported=encodeDocxWord(bytes);try{const result=imported.doc.getArray<Y.Text>('paragraphs');expect(result.toArray().map(value=>value.getAttribute('firstLineIndent'))).toEqual([720,-720,0,-123]);expect(result.toArray().map(value=>value.getAttribute('rightIndent'))).toEqual([360,2880,0,123]);expect(result.get(0).toDelta()).toEqual(p.get(0).toDelta());}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('特殊缩进继承按模式覆盖，悬挂优先，显式零值清除继承',async()=>{
 const {wordLayoutReader}=await import('../packages/frontend/apps/web/server/office-word-layout');
 const read=wordLayoutReader(new TextEncoder().encode('<w:styles><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:ind w:right="500" w:hanging="720"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:pPr><w:ind w:firstLine="360"/></w:pPr></w:style></w:styles>'));
 expect(read('<w:pPr><w:pStyle w:val="Child"/></w:pPr>')).toEqual({rightIndent:500,firstLineIndent:360});
 expect(read('<w:pPr><w:pStyle w:val="Base"/><w:ind w:right="0" w:firstLine="0"/></w:pPr>')).toEqual({rightIndent:0,firstLineIndent:0});
 expect(read('<w:pPr><w:ind w:firstLine="720" w:hanging="360"/></w:pPr>')).toEqual({firstLineIndent:-360});expect(read('<w:pPr><w:ind w:hanging="0"/></w:pPr>')).toEqual({firstLineIndent:0});
 for(const attrs of ['w:hanging="2881"','w:firstLine="-1"','w:right="1.5"','w:firstLineChars="200"'])expect(()=>read(`<w:pPr><w:ind ${attrs}/></w:pPr>`)).toThrow('缩进');
});
it('更新拒绝无效的右缩进与首行属性',async()=>{
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const [key,value]of [['rightIndent',-1],['rightIndent',2881],['rightIndent','720'],['firstLineIndent',-2881],['firstLineIndent',2881],['firstLineIndent',1.5],['firstLineIndent',{}]] as [string,unknown][]){const doc=snapshot(['文']);try{doc.getArray<Y.Text>('paragraphs').get(0).setAttribute(key,value);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('段落属性');}finally{doc.destroy();}}
});

it('整篇页面尺寸、方向、页边距及段前分页往返，旧模型仍可导出',async()=>{
 const {DEFAULT_WORD_PAGE}=await import('../packages/common/nbstore/src/office-word-page');
 const page={width:15840,height:12240,top:720,right:900,bottom:1080,left:1234},doc=snapshot(['第一页','第二页']);
 try{
  doc.getMap('pageSettings').set('layout',page);doc.getArray<Y.Text>('paragraphs').get(0).setAttribute('pageBreakBefore',false);doc.getArray<Y.Text>('paragraphs').get(1).setAttribute('pageBreakBefore',true);
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),imported=encodeDocxWord(bytes);try{expect(imported.doc.getMap('pageSettings').get('layout')).toEqual(page);expect(imported.doc.getArray<Y.Text>('paragraphs').toArray().map(p=>p.getAttribute('pageBreakBefore'))).toEqual([false,true]);}finally{imported.doc.destroy();}
  doc.getMap('pageSettings').delete('layout');const legacy=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(doc)));try{expect(legacy.doc.getMap('pageSettings').get('layout')).toEqual(DEFAULT_WORD_PAGE);}finally{legacy.doc.destroy();}
 }finally{doc.destroy();}
});
it('拒绝无效页面模型、隐藏根序列、未知字段及非布尔分页',async()=>{
 const {DEFAULT_WORD_PAGE}=await import('../packages/common/nbstore/src/office-word-page');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const page of [{...DEFAULT_WORD_PAGE,width:'11906'},{...DEFAULT_WORD_PAGE,left:7200,right:7200},{...DEFAULT_WORD_PAGE,width:31681},{...DEFAULT_WORD_PAGE,top:-1},{...DEFAULT_WORD_PAGE,hidden:1}]){const doc=snapshot(['文']);try{doc.getMap('pageSettings').set('layout',page);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('页面设置');}finally{doc.destroy();}}
 const {validWordPage}=await import('../packages/common/nbstore/src/office-word-page');expect(validWordPage(Object.assign(Object.create(DEFAULT_WORD_PAGE),{a:1,b:2,c:3,d:4,e:5,f:6}))).toBe(false);
 const hidden=snapshot(['文']);try{hidden.getText('pageSettings').insert(0,'hidden');expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(hidden))).toThrow('页面设置');}finally{hidden.destroy();}
 const invalid=snapshot(['文']);try{invalid.getArray<Y.Text>('paragraphs').get(0).setAttribute('pageBreakBefore','true');expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(invalid))).toThrow('段落属性');}finally{invalid.destroy();}
});
it('拒绝多分节、装订线、多栏和段内分页，段前分页可继承并显式关闭',async()=>{
 const {readWordPage}=await import('../packages/frontend/apps/web/server/office-word-page');const {wordLayoutReader}=await import('../packages/frontend/apps/web/server/office-word-layout');
 expect(()=>readWordPage('<w:document><w:sectPr/><w:sectPr/></w:document>')).toThrow('多分节');expect(()=>readWordPage('<w:document><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:document>')).toThrow('多分节');
 for(const xml of ['<w:pgMar w:gutter="100"/>','<w:cols w:num="2"/>','<w:cols><w:col w:w="1000"/><w:col w:w="1000"/></w:cols>'])expect(()=>readWordPage('<w:document><w:sectPr>'+xml+'</w:sectPr></w:document>')).toThrow('暂不支持');
 const {zipSync,strToU8}=await import('fflate');expect(()=>encodeDocxWord(zipSync({'word/document.xml':strToU8('<w:document><w:p><w:r><w:br w:type="page"/></w:r></w:p></w:document>')}))).toThrow('段内分页');
 const read=wordLayoutReader(new TextEncoder().encode('<w:styles><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:pageBreakBefore/></w:pPr></w:style></w:styles>'));expect(read('<w:pPr><w:pStyle w:val="Base"/></w:pPr>')).toEqual({pageBreakBefore:true});expect(read('<w:pPr><w:pStyle w:val="Base"/><w:pageBreakBefore w:val="0"/></w:pPr>')).toEqual({pageBreakBefore:false});
});

it('字号、RGB 和自动色在普通文字和链接中往返，保留现有样式',()=>{
 const doc=snapshot(['半磅字号自动色']);
 try{
  const p=doc.getArray<Y.Text>('paragraphs').get(0);p.format(0,3,{fontSize:13.5,color:'#123ABC',bold:true,href:'https://example.com/'});p.format(3,4,{fontSize:200,color:'auto',italic:true});
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),xml=readOfficeXml(readOfficePackage(bytes)['word/document.xml']);expect(xml).toContain('<w:sz w:val="27"/>');expect(xml).toContain('<w:color w:val="123ABC"/>');expect(xml).toContain('<w:color w:val="auto"/>');
  const imported=encodeDocxWord(bytes);try{expect(imported.doc.getArray<Y.Text>('paragraphs').get(0).toDelta()).toEqual(p.toDelta());}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('导入拒绝主题色、非法 RGB 和超范围字号',async()=>{
 const {zipSync,strToU8}=await import('fflate');
 for(const properties of ['<w:sz w:val="0"/>','<w:sz w:val="401"/>','<w:sz w:val="24.5"/>','<w:color w:val="red"/>','<w:color w:val="FF0000" w:themeColor="accent1"/>'])expect(()=>encodeDocxWord(zipSync({'word/document.xml':strToU8(`<w:document><w:p><w:r><w:rPr>${properties}</w:rPr><w:t>文字</w:t></w:r></w:p></w:document>`)}))).toThrow();
});
it('服务端拒绝非半磅字号、CSS 注入颜色及错误属性类型',async()=>{
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const attributes of [{fontSize:0},{fontSize:200.5},{fontSize:12.25},{fontSize:'12'},{color:'red'},{color:'url(https://example.com/)'},{color:{hex:'#000000'}}]){const doc=snapshot(['文字']);try{doc.getArray<Y.Text>('paragraphs').get(0).format(0,2,attributes);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('富文本');}finally{doc.destroy();}}
});

it('DOCX 段落与基础表格保持顺序，单元格文字、空值、字号和链接往返',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');
 const doc=snapshot(['表前']);
 try{
  const table=createWordTable(2,2);doc.getArray('paragraphs').push([table,new Y.Text('表后')]);const rows=wordTableRows(table);rows.get(0).get(0).insert(0,'中文😀');rows.get(0).get(0).format(0,4,{fontSize:13.5,bold:true,href:'https://example.com/'});rows.get(1).get(1).insert(0,'末行\t换列\n换行');
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),xml=readOfficeXml(readOfficePackage(bytes)['word/document.xml']);expect(xml.indexOf('表前')).toBeLessThan(xml.indexOf('<w:tbl>'));expect(xml.indexOf('</w:tbl>')).toBeLessThan(xml.indexOf('表后'));expect(xml).toContain('<w:tblGrid>');
  const imported=encodeDocxWord(bytes);try{const blocks=imported.doc.getArray('paragraphs');expect(blocks.length).toBe(3);expect((blocks.get(0) as Y.Text).toString()).toBe('表前');expect((blocks.get(2) as Y.Text).toString()).toBe('表后');const result=wordTableRows(blocks.get(1) as Y.Map<unknown>);expect(result.length).toBe(2);expect(result.get(0).get(0).toDelta()).toEqual(rows.get(0).get(0).toDelta());expect(result.get(0).get(1).toString()).toBe('');expect(result.get(1).get(1).toString()).toBe('末行\t换列\n换行');}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
it('DOCX 表格拒绝不完整合并、嵌套、格式不一致的多段落、空单元格结构及不规则网格',async()=>{
 const {zipSync,strToU8}=await import('fflate');
 for(const inner of [
  '<w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr>',
  '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>',
  '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>隐藏</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>',
  '<w:tr><w:tc><w:tcPr><w:hMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr>',
  '<w:tr><w:tc><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p><w:p/></w:tc></w:tr>',
  '<w:tr><w:tc/></w:tr>',
  '<w:tr><w:tc><w:p/></w:tc></w:tr><w:tr/>',
  '<w:tr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>',
 ])expect(()=>encodeDocxWord(zipSync({'word/document.xml':strToU8(`<w:document><w:body><w:tbl>${inner}</w:tbl></w:body></w:document>`)}))).toThrow();
});
it('服务端拒绝非法表格列数、额外字段、行结构和非文字单元格',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 for(const mutation of [(table:Y.Map<unknown>)=>table.set('columns',3),(table:Y.Map<unknown>)=>table.set('hidden','x'),(table:Y.Map<unknown>)=>wordTableRows(table).get(0).delete(0,1),(table:Y.Map<unknown>)=>{const row=wordTableRows(table).get(0);row.delete(0,1);(row as Y.Array<unknown>).insert(0,[new Y.Map()]);}]){
  const doc=snapshot([]);try{const table=createWordTable(2,2);doc.getArray('paragraphs').push([table]);mutation(table);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('Word');}finally{doc.destroy();}
 }
});
it('空表格允许协作恢复，导出不会凭空补行',async()=>{
 const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');const doc=snapshot([]);
 try{const table=createWordTable(2,0);doc.getArray('paragraphs').push([table]);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();expect(()=>exportDocxWord(Y.encodeStateAsUpdate(doc))).toThrow('空表格');const row=new Y.Array<Y.Text>();row.push([new Y.Text('恢复'),new Y.Text()]);wordTableRows(table).push([row]);expect(()=>exportDocxWord(Y.encodeStateAsUpdate(doc))).not.toThrow();}finally{doc.destroy();}
});

it('备份模型验证保留 Word 表格，清除旧表格回执，拒绝缺依赖与图片引用',async()=>{
 const {validateOfficeBackupModel}=await import('../packages/frontend/apps/web/server/office-backup-model');const {createWordTable,wordTableRows}=await import('../packages/common/nbstore/src/office-word-table');const word=snapshot(['正文']),sheet=new Y.Doc(),slide=new Y.Doc(),partial=new Y.Doc();
 try{const table=createWordTable(2,2);word.getArray('paragraphs').push([table]);wordTableRows(table).get(0).get(0).insert(0,'表格😀');const validated=validateOfficeBackupModel('word',Y.encodeStateAsUpdate(word));expect(exportDocxWord(validated)).toBeInstanceOf(Uint8Array);
  const row=new Y.Array<string>();row.push(['=1+2','中文']);sheet.getArray('sheet').push([row]);sheet.getMap('sheetOperations').set('old','a'.repeat(64));const restored=new Y.Doc();try{Y.applyUpdate(restored,validateOfficeBackupModel('spreadsheet',Y.encodeStateAsUpdate(sheet)));expect(restored.getMap('sheetOperations').size).toBe(0);expect((restored.getArray('sheet').get(0) as Y.Array<string>).toArray()).toEqual(['=1+2','中文']);expect(sheet.getMap('sheetOperations').size).toBe(1);}finally{restored.destroy();}
  const image=new Y.Map<unknown>();for(const[k,v]of Object.entries({type:'image',x:0,y:0,width:1,height:1}))image.set(k,v);slide.getArray('objects').push([image]);expect(()=>validateOfficeBackupModel('presentation',Y.encodeStateAsUpdate(slide))).toThrow('图片引用');
  partial.getText('text').insert(0,'原文');const vector=Y.encodeStateVector(partial);partial.getText('text').insert(2,'后续');expect(()=>validateOfficeBackupModel('text',Y.encodeStateAsUpdate(partial,vector))).toThrow();
 }finally{word.destroy();sheet.destroy();slide.destroy();partial.destroy();}
});

it('文字格式按文档默认、段落样式链与字符样式链继承，直接格式覆盖，正文基线不写入且原生副本往返',async()=>{
 const {zipSync,strToU8}=await import('fflate');
 const styles='<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/><w:color w:val="111111"/></w:rPr></w:rPrDefault></w:docDefaults>'
  +'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
  +'<w:style w:type="paragraph" w:styleId="Quote"><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="ff0000"/><w:rPrChange w:id="1"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr></w:style>'
  +'<w:style w:type="paragraph" w:styleId="BigQuote"><w:basedOn w:val="Quote"/><w:rPr><w:sz w:val="36"/><w:u w:val="single"/></w:rPr></w:style>'
  +'<w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:themeColor="accent1" w:val="2F5496"/></w:rPr></w:style>'
  +'<w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="Emph"><w:basedOn w:val="Strong"/><w:rPr><w:i/><w:sz w:val="999"/></w:rPr></w:style></w:styles>';
 const run=(text:string,rPr='')=>`<w:r>${rPr?'<w:rPr>'+rPr+'</w:rPr>':''}<w:t>${text}</w:t></w:r>`,paragraph=(style:string,runs:string)=>`<w:p>${style?`<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`:''}${runs}</w:p>`;
 const body=[paragraph('',run('正文')+run('强调','<w:rStyle w:val="Strong"/>')),paragraph('BigQuote',run('引用')+run('取消','<w:i w:val="0"/><w:u w:val="none"/>')+run('切换','<w:rStyle w:val="Emph"/>')),paragraph('Heading1',run('标题'))].join('');
 const parsed=encodeDocxWord(zipSync({'word/styles.xml':strToU8(styles),'word/document.xml':strToU8('<w:document><w:body>'+body+'</w:body></w:document>')}));
 try{
  const [normal,quote,heading]=parsed.doc.getArray<Y.Text>('paragraphs').toArray();
  expect(normal.toDelta()).toEqual([{insert:'正文'},{insert:'强调',attributes:{bold:true}}]);
  // Italic toggles off when a character style based on Strong adds italic on top of an italic paragraph style; invalid inherited sizes are ignored.
  expect(quote.toDelta()).toEqual([{insert:'引用',attributes:{italic:true,underline:true,fontSize:18,color:'#FF0000'}},{insert:'取消',attributes:{italic:false,underline:false,fontSize:18,color:'#FF0000'}},{insert:'切换',attributes:{bold:true,underline:true,fontSize:18,color:'#FF0000'}}]);
  expect(heading.getAttribute('heading')).toBe(1);expect(heading.toDelta()).toEqual([{insert:'标题'}]);
  const again=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(parsed.doc)));try{expect(again.doc.getArray<Y.Text>('paragraphs').toArray().map(text=>text.toDelta())).toEqual([normal,quote,heading].map(text=>text.toDelta()));}finally{again.doc.destroy();}
 }finally{parsed.doc.destroy();}
 expect(()=>encodeDocxWord(zipSync({'word/styles.xml':strToU8('<w:styles><w:style w:type="character" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="character" w:styleId="B"><w:basedOn w:val="A"/></w:style></w:styles>'),'word/document.xml':strToU8('<w:document><w:body>'+paragraph('',run('x','<w:rStyle w:val="A"/>'))+'</w:body></w:document>')}))).toThrow('样式继承');
});

it('Word 合并单元格与单元格多段落：导入 gridSpan/vMerge、原生副本往返、编辑合并拆分及行列操作保持内容',async()=>{
 const {zipSync,strToU8}=await import('fflate');const {wordTableRows,wordTableLayout,WORD_CELL_PARAGRAPH}=await import('../packages/common/nbstore/src/office-word-table');
 const {mergeWordTableCell,splitWordTableCell,deleteWordTableRow,insertWordTableRow,deleteWordTableColumn,insertWordTableColumn}=await import('../packages/frontend/core/src/office-word-table');const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 const p=(text:string,ppr='')=>`<w:p>${ppr?'<w:pPr>'+ppr+'</w:pPr>':''}${text?`<w:r><w:t>${text}</w:t></w:r>`:''}</w:p>`,tc=(content:string,tcPr='')=>`<w:tc>${tcPr?'<w:tcPr>'+tcPr+'</w:tcPr>':''}${content}</w:tc>`;
 const table='<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>'
  +'<w:tr>'+tc(p('标题','<w:jc w:val="center"/>')+p('第二段','<w:jc w:val="center"/>'),'<w:gridSpan w:val="2"/>')+tc(p('右上'))+'</w:tr>'
  +'<w:tr>'+tc(p('左侧'),'<w:vMerge w:val="restart"/>')+tc(p('中'))+tc(p('右'))+'</w:tr>'
  +'<w:tr>'+tc(p(''),'<w:vMerge/>')+tc(p('下中'))+tc(p('下右'))+'</w:tr></w:tbl>';
 const parsed=encodeDocxWord(zipSync({'word/document.xml':strToU8('<w:document><w:body>'+table+'</w:body></w:document>')}));
 try{
  const block=parsed.doc.getArray('paragraphs').get(0) as Y.Map<unknown>,rows=wordTableRows(block),layout=wordTableLayout(block),title=rows.get(0).get(0),left=rows.get(1).get(0);
  expect(title.toString()).toBe('标题'+WORD_CELL_PARAGRAPH+'第二段');expect(title.getAttribute('align')).toBe('center');expect(layout.anchors.get(title)).toMatchObject({colSpan:2,rowSpan:1});expect(layout.anchors.get(left)).toMatchObject({rowSpan:2});expect(layout.anchors.has(rows.get(0).get(1))).toBe(false);expect(layout.anchors.has(rows.get(2).get(0))).toBe(false);
  const xml=readOfficeXml(readOfficePackage(exportDocxWord(Y.encodeStateAsUpdate(parsed.doc)))['word/document.xml']);expect(xml).toContain('<w:gridSpan w:val="2"/>');expect(xml).toContain('<w:vMerge w:val="restart"/>');expect(xml.match(/<w:tc>/g)).toHaveLength(8);expect(xml.match(/<w:jc w:val="center"\/>/g)).toHaveLength(2);
  const again=encodeDocxWord(exportDocxWord(Y.encodeStateAsUpdate(parsed.doc)));try{const copy=again.doc.getArray('paragraphs').get(0) as Y.Map<unknown>;expect(wordTableRows(copy).toJSON()).toEqual(rows.toJSON());expect([...wordTableLayout(copy).anchors.values()]).toEqual([...layout.anchors.values()]);}finally{again.doc.destroy();}
  // Editing: merging collects text as cell paragraphs; split reveals empty cells; row/column edits keep the anchor's content.
  const middle=rows.get(1).get(1);mergeWordTableCell(block,middle,'right');expect(middle.toString()).toBe('中'+WORD_CELL_PARAGRAPH+'右');expect(rows.get(1).get(2).toString()).toBe('');expect(()=>mergeWordTableCell(block,middle,'right')).toThrow('右侧没有');
  mergeWordTableCell(block,middle,'down');expect(wordTableLayout(block).anchors.get(middle)).toMatchObject({colSpan:2,rowSpan:2});expect(middle.toString().split(WORD_CELL_PARAGRAPH)).toEqual(['中','右','下中','下右']);
  expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(parsed.doc))).not.toThrow();
  insertWordTableRow(block,rows.get(1),'after');expect(wordTableLayout(block).anchors.get(middle)?.rowSpan).toBe(3);expect(wordTableLayout(block).anchors.get(left)?.rowSpan).toBe(3);
  deleteWordTableRow(block,rows.get(1));const moved=rows.get(1).get(0),movedMiddle=rows.get(1).get(1);expect(moved.toString()).toBe('左侧');expect(wordTableLayout(block).anchors.get(moved)?.rowSpan).toBe(2);expect(movedMiddle.toString().split(WORD_CELL_PARAGRAPH)).toHaveLength(4);expect(wordTableLayout(block).anchors.get(movedMiddle)).toMatchObject({colSpan:2,rowSpan:2});
  insertWordTableColumn(block,movedMiddle,'after');expect(wordTableLayout(block).anchors.get(movedMiddle)?.colSpan).toBe(2);deleteWordTableColumn(block,movedMiddle);const shifted=rows.get(1).get(1);expect(shifted.toString().split(WORD_CELL_PARAGRAPH)).toHaveLength(4);expect(wordTableLayout(block).anchors.get(shifted)).toMatchObject({colSpan:1,rowSpan:2});
  splitWordTableCell(block,shifted);expect(wordTableLayout(block).anchors.get(rows.get(2).get(1))).toBeDefined();expect(()=>splitWordTableCell(block,shifted)).toThrow('未合并');
  // A stale span that would cover text degrades to separate cells instead of hiding it.
  rows.get(2).get(1).insert(0,'并发输入');shifted.setAttribute('rowSpan',2);expect(wordTableLayout(block).anchors.get(shifted)).toMatchObject({rowSpan:1});expect(wordTableLayout(block).anchors.has(rows.get(2).get(1))).toBe(true);
 }finally{parsed.doc.destroy();}
 for(const bad of [(doc:Y.Doc)=>(doc.getArray('paragraphs').get(1) as Y.Text).setAttribute('colSpan',2),(doc:Y.Doc)=>(doc.getArray('paragraphs').get(1) as Y.Text).insert(0,WORD_CELL_PARAGRAPH),(doc:Y.Doc)=>wordTableRows(doc.getArray('paragraphs').get(0) as Y.Map<unknown>).get(0).get(0).setAttribute('rowSpan',1.5),(doc:Y.Doc)=>wordTableRows(doc.getArray('paragraphs').get(0) as Y.Map<unknown>).get(0).get(0).insert(0,WORD_CELL_PARAGRAPH.repeat(100))]){
  const {createWordTable}=await import('../packages/common/nbstore/src/office-word-table');const doc=snapshot([]);try{doc.getArray('paragraphs').push([createWordTable(2,2),new Y.Text('正文')]);bad(doc);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('Word');}finally{doc.destroy();}
 }
});

it('删除线、上下标、突出显示、字体及行距/段前段后经 DOCX 往返，样式继承删除线，非法值被拒绝',async()=>{
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');const {zipSync,strToU8}=await import('fflate');
 const doc=snapshot([]);try{const text=new Y.Text();text.insert(0,'删除',{strike:true});text.insert(2,'上',{vertAlign:'superscript',fontSize:12});text.insert(3,'下',{vertAlign:'subscript'});text.insert(4,'亮',{highlight:'yellow'});text.insert(5,'宋体',{fontFamily:'宋体',bold:true});text.setAttribute('lineSpacing',1.5);text.setAttribute('spaceBefore',240);text.setAttribute('spaceAfter',120);doc.getArray('paragraphs').push([text]);
  expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),xml=readOfficeXml(readOfficePackage(bytes)['word/document.xml']);
  expect(xml).toContain('<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>');expect(xml).toContain('<w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/><w:b w:val="1"/>');expect(xml).toContain('<w:highlight w:val="yellow"/>');expect(xml).toContain('<w:vertAlign w:val="superscript"/>');
  const again=encodeDocxWord(bytes);try{const copy=again.doc.getArray<Y.Text>('paragraphs').get(0);expect(copy.toDelta()).toEqual(text.toDelta());expect(copy.getAttributes()).toEqual(text.getAttributes());}finally{again.doc.destroy();}
  for(const [key,value]of [['lineSpacing',9],['lineSpacing',1.234],['spaceBefore',-1],['spaceAfter',2000]] as const){const copy=new Y.Doc();try{Y.applyUpdate(copy,Y.encodeStateAsUpdate(doc));copy.getArray<Y.Text>('paragraphs').get(0).setAttribute(key,value);expect(()=>validateWordUpdate(null,Y.encodeStateAsUpdate(copy))).toThrow('段落属性');}finally{copy.destroy();}}
 }finally{doc.destroy();}
 const styles='<w:styles><w:style w:type="paragraph" w:styleId="Old"><w:pPr><w:spacing w:before="120" w:after="0" w:line="480" w:lineRule="auto"/></w:pPr><w:rPr><w:strike/></w:rPr></w:style></w:styles>',body='<w:p><w:pPr><w:pStyle w:val="Old"/></w:pPr><w:r><w:t>旧版</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:line="300" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:highlight w:val="none"/><w:vertAlign w:val="baseline"/></w:rPr><w:t>默认</w:t></w:r></w:p>';
 const parsed=encodeDocxWord(zipSync({'word/styles.xml':strToU8(styles),'word/document.xml':strToU8('<w:document><w:body>'+body+'</w:body></w:document>')}));try{const [old,plain]=parsed.doc.getArray<Y.Text>('paragraphs').toArray();expect(old.toDelta()).toEqual([{insert:'旧版',attributes:{strike:true}}]);expect(old.getAttributes()).toEqual({lineSpacing:2,spaceBefore:120,spaceAfter:0});expect(plain.toDelta()).toEqual([{insert:'默认'}]);expect(plain.getAttributes()).toEqual({});}finally{parsed.doc.destroy();}
 expect(()=>encodeDocxWord(zipSync({'word/document.xml':strToU8('<w:document><w:body><w:p><w:r><w:rPr><w:highlight w:val="orange"/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>')}))).toThrow('突出显示');
});

it('文件评论导出为 Word 原生批注：锚定文字与单元格段落、回复成线程、已解决标记、失效定位置于文首，正文往返不变',async()=>{
 const {createWordTable,wordTableRows,WORD_CELL_PARAGRAPH}=await import('../packages/common/nbstore/src/office-word-table');
 const {captureOfficeTextAnchor}=await import('../packages/common/nbstore/src/office-comment-anchor');
 const doc=snapshot(['甲乙丙丁戊','第二段']),blocks=doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs'),table=createWordTable(1,1);blocks.push([table]);
 const first=blocks.get(0) as Y.Text,second=blocks.get(1) as Y.Text,cell=wordTableRows(table).get(0).get(0);cell.insert(0,'格一'+WORD_CELL_PARAGRAPH+'格二');
 const comments=doc.getMap('officeComments'),message=(id:string,author:string,body:string,created:number)=>({id,author:'u'+author.length,authorName:author,body,created});
 comments.set('t1',{id:'t1',resolved:false,updated:3,anchor:captureOfficeTextAnchor(first,1,first,3,'乙丙'),messages:[message('m1','王小明','请核对 <数字> & 单位',1),message('m2','Lisa Chen','已核对\n第二行',2)]});
 comments.set('t2',{id:'t2',resolved:true,updated:5,anchor:captureOfficeTextAnchor(cell,1,cell,4,'一 格'),messages:[message('m3','王小明','跨格段',4)]});
 comments.set('t3',{id:'t3',resolved:false,updated:6,messages:[message('m4','王小明','整篇意见',6)]});
 comments.set('t4',{id:'t4',resolved:false,updated:7,anchor:captureOfficeTextAnchor(second,0,second,1,'第'),messages:[message('m5','王小明','原文已删',7)]});
 second.delete(0,second.length);
 try{
  const bytes=exportDocxWord(Y.encodeStateAsUpdate(doc)),files=readOfficePackage(bytes),xml=readOfficeXml(files['word/document.xml']),notes=readOfficeXml(files['word/comments.xml']),extended=readOfficeXml(files['word/commentsExtended.xml']);
  expect(readOfficeXml(files['[Content_Types].xml'])).toContain('/word/comments.xml');expect(readOfficeXml(files['word/_rels/document.xml.rels'])).toContain('Target="comments.xml"');
  expect([...notes.matchAll(/<w:comment /g)]).toHaveLength(5);expect(notes).toContain('w:author="Lisa Chen"');expect(notes).toContain('请核对 &lt;数字&gt; &amp; 单位');expect(notes).toContain('w:date="1970-01-01T00:00:00Z"');
  // Root comment of t1 wraps exactly "乙丙"; its reply shares the range and is threaded under it.
  expect(xml).toMatch(/甲<\/w:t><\/w:r><w:commentRangeStart w:id="\d+"\/><w:commentRangeStart w:id="\d+"\/><w:r><w:t xml:space="preserve">乙丙<\/w:t><\/w:r><w:commentRangeEnd/);
  const ex=[...extended.matchAll(/<w15:commentEx ([^>]*)\/>/g)].map(m=>m[1]);expect(ex).toHaveLength(5);expect(ex.filter(e=>e.includes('paraIdParent'))).toHaveLength(1);expect(ex.filter(e=>e.includes('w15:done="1"'))).toHaveLength(1);
  // Every range start has a matching end and a reference.
  for(const id of [0,1,2,3,4]){expect(xml).toContain(`<w:commentRangeStart w:id="${id}"/>`);expect(xml).toContain(`<w:commentRangeEnd w:id="${id}"/>`);expect(xml).toContain(`<w:commentReference w:id="${id}"/>`);}
  // The cell comment starts in the cell's first paragraph and ends in its second.
  const cellXml=xml.slice(xml.indexOf('<w:tbl>')),split=cellXml.indexOf('</w:p><w:p>');expect(cellXml).toMatch(/<w:t xml:space="preserve">格<\/w:t><\/w:r><w:commentRangeStart w:id="2"\/><w:r><w:t xml:space="preserve">一<\/w:t>/);expect(cellXml.indexOf('<w:commentRangeStart w:id="2"/>')).toBeLessThan(split);expect(cellXml.indexOf('<w:commentRangeEnd w:id="2"/>')).toBeGreaterThan(split);expect(cellXml).toContain('<w:commentRangeEnd w:id="2"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="2"/></w:r><w:r><w:t xml:space="preserve">二</w:t>');
  // Floating comments (no anchor, or anchor deleted) sit at the start of the first paragraph.
  expect(xml.indexOf('<w:commentRangeStart w:id="3"/>')).toBeLessThan(xml.indexOf('甲'));expect(xml.indexOf('<w:commentRangeStart w:id="4"/>')).toBeLessThan(xml.indexOf('甲'));
  const imported=encodeDocxWord(bytes);try{const texts=imported.doc.getArray<Y.Text|Y.Map<unknown>>('paragraphs').toArray();expect((texts[0] as Y.Text).toString()).toBe('甲乙丙丁戊');expect(wordTableRows(texts[2] as Y.Map<unknown>).get(0).get(0).toString()).toBe('格一'+WORD_CELL_PARAGRAPH+'格二');}finally{imported.doc.destroy();}
 }finally{doc.destroy();}
});
