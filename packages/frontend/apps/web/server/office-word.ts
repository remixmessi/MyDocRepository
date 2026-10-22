import {validateOfficeComments} from '../../../../common/nbstore/src/office-comments';
import {packageImageReader} from './office-media';
import {isWordImage,createWordImage,validateWordImage,OFFICE_IMAGE_COUNT} from '../../../../common/nbstore/src/office-image';
import {createWordTable,wordTableRows,validateWordTable,wordTableLayout,isWordTable,wordBlockTexts,WORD_TABLE_LIMITS,WORD_CELL_PARAGRAPH,type WordBlock} from '../../../../common/nbstore/src/office-word-table';
import {wordItemIdentity,wordTableStructure} from '../../../../common/nbstore/src/office-word-table-edit';
import {readWordTable} from './office-word-table';
import {validWordFontSize,validWordColor,validWordHighlight,validWordVertAlign,validWordFontFamily} from '../../../../common/nbstore/src/office-word-font';
import {validWordPage} from '../../../../common/nbstore/src/office-word-page';
import {validWordHeaderFooter} from '../../../../common/nbstore/src/office-word-header';
import {validWordRevisionMark,WORD_REVISION_KEYS} from '../../../../common/nbstore/src/office-word-revision';
import {readWordHeaderFooter} from './office-word-headers';
import {readWordPage} from './office-word-page';
import {validWordAlignment,validWordLeftIndent,validWordRightIndent,validWordFirstLineIndent,validWordLineSpacing,validWordParagraphSpace} from '../../../../common/nbstore/src/office-word-layout';
import {wordLayoutReader} from './office-word-layout';
import {validWordLink} from '../../../../common/nbstore/src/office-word-link';
import {decodeWordXml,wordLinkReader,wordRuns,type WordRunRevision} from './office-word-link';
import {validWordList,type WordList} from '../../../../common/nbstore/src/office-word-list';
import {wordListReader} from './office-word-list';
import * as Y from 'yjs';
import {wordHeadingReader} from './office-word-heading';
import {wordRunStyleReader,wordRunProperties} from './office-word-run-style';
import {readOfficePackage,readOfficeXml} from './office-package';
export const MAX_WORD_PARAGRAPHS=100_000;
const MARKS=['bold','italic','underline','strike'];
function runAttributes(xml:string){
 const properties=wordRunProperties(xml);
 const attributes:Record<string,boolean|string|number>={};
 for(const [tag,key]of [['b','bold'],['i','italic'],['u','underline'],['strike','strike']]){
  const property=new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*?)(?:/>|>[\\s\\S]*?</w:${tag}>)`).exec(properties);
  if(!property)continue;
  const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(property[1])?.[2];
  if(tag==='u'){if(value==='single')attributes[key]=true;else if(value==='none')attributes[key]=false;}
  else if(value===undefined||['1','true','on'].includes(value))attributes[key]=true;
  else if(['0','false','off'].includes(value))attributes[key]=false;
 }
 const size=/<w:sz(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:sz>)/.exec(properties);
 if(size){const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(size[1])?.[2],points=Number(value)/2;if(value===undefined||!/^\d+$/.test(value)||!validWordFontSize(points))throw new Error('Word 字号须在 1–200 磅且为半磅步进');attributes.fontSize=points;}
 const color=/<w:color(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:color>)/.exec(properties);
 if(color){if(/\bw:theme(?:Color|Tint|Shade)\s*=/.test(color[1]))throw new Error('暂不支持主题文字颜色，请转换为 RGB 颜色');const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(color[1])?.[2],normalized=value==='auto'?'auto':'#'+(value??'').toUpperCase();if(!validWordColor(normalized))throw new Error('Word 文字颜色无效');attributes.color=normalized;}
 const vert=/<w:vertAlign(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:vertAlign>)/.exec(properties);
 if(vert){const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(vert[1])?.[2];if(validWordVertAlign(value))attributes.vertAlign=value;else if(value!=='baseline')throw new Error('Word 上下标设置无效');}
 const highlight=/<w:highlight(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:highlight>)/.exec(properties);
 if(highlight){const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(highlight[1])?.[2];if(validWordHighlight(value))attributes.highlight=value;else if(value!=='none')throw new Error('暂不支持该 Word 突出显示颜色');}
 // Explicit font names only; theme font references are left to the document default.
 const fonts=/<w:rFonts(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:rFonts>)/.exec(properties);
 if(fonts){const pick=(name:string)=>{const value=new RegExp(`\\bw:${name}\\s*=\\s*(["'])(.*?)\\1`).exec(fonts[1])?.[2];return value===undefined?undefined:decodeWordXml(value);},family=pick('eastAsia')??pick('ascii')??pick('hAnsi');if(family!==undefined&&validWordFontFamily(family))attributes.fontFamily=family;}
 return attributes;
}
export function encodeDocxWord(bytes:Uint8Array){
 const files=readOfficePackage(bytes),data=files['word/document.xml'];if(!data)throw new Error('DOCX 缺少正文');
 const xml=readOfficeXml(data),headingFor=wordHeadingReader(files['word/styles.xml']),listFor=wordListReader(files['word/numbering.xml'],files['word/styles.xml']),linkFor=wordLinkReader(files['word/_rels/document.xml.rels']),layoutFor=wordLayoutReader(files['word/styles.xml']),runStyleFor=wordRunStyleReader(files['word/styles.xml']),doc=new Y.Doc(),paragraphs=doc.getArray<WordBlock>('paragraphs');
 try{
  const images=packageImageReader(files,'word/document.xml');
  const appendImages=(source:string)=>{for(const drawing of source.matchAll(/<w:drawing\b[\s\S]*?<\/w:drawing>/g)){const value=drawing[0],embed=/\br:embed\s*=\s*(["'])(.*?)\1/.exec(value)?.[2];if(!embed)throw new Error('DOCX 图片缺少内嵌关系');const extent=/<wp:extent\b([^>]*?)\/?>/.exec(value)?.[1]??'',width=Number(/\bcx=["'](\d+)["']/.exec(extent)?.[1])/9525,height=Number(/\bcy=["'](\d+)["']/.exec(extent)?.[1])/9525;const alt=/<wp:docPr\b[^>]*\bdescr=(["'])(.*?)\1/.exec(value)?.[2]??'图片';paragraphs.push([createWordImage(images.read(embed),width,height,decodeWordXml(alt))]);}};
  const page=readWordPage(xml);if(page)doc.getMap('pageSettings').set('layout',page);
  const headerFooter=readWordHeaderFooter(files,xml);if(headerFooter)doc.getMap('pageSettings').set('headerFooter',headerFooter);
  // Revisions keep Word's author and date; the id is made unique per kind so ins/del of one edit stay distinct.
  const revisionMark=(revision:WordRunRevision)=>{const date=Date.parse(revision.date??''),name=revision.author.trim().slice(0,160)||'Word 作者';return {id:('docx-'+revision.kind+'-'+revision.id).replace(/[^a-zA-Z0-9-]/g,'-').slice(0,80),author:'docx-author',name,date:Number.isFinite(date)&&date>0?date:1};};
  const settingsXml=files['word/settings.xml']?readOfficeXml(files['word/settings.xml']):'',tracked=/<w:trackRevisions(?=[\s/>])([^>]*)\/?>/.exec(settingsXml);if(tracked&&!/w:val\s*=\s*(["'])(?:0|false|off)\1/.test(tracked[1]))doc.getMap('pageSettings').set('trackChanges',true);
  if(/<w:br\b[^>]*w:type\s*=\s*(["'])page\1/.test(xml))throw new Error('暂不支持段内分页符，请改用段前分页');
  const paragraphAttributes=(xml:string)=>{const heading=headingFor(xml),list=listFor(xml);return {...(heading?{heading}:{}),...(list?{list}:{}),...layoutFor(xml)} as Record<string,unknown>;};
  const fill=(text:Y.Text,xml:string,attributes=true,cell=false)=>{
   const heading=headingFor(xml);if(attributes)for(const [key,value]of Object.entries(paragraphAttributes(xml)))text.setAttribute(key,value);
   const paragraphStyle=/<w:pStyle\b[^>]*\bw:val\s*=\s*(["'])(.*?)\1/.exec(/<w:pPr(?:\s[^>]*)?>([\s\S]*?)<\/w:pPr>/.exec(xml.replace(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/g,''))?.[1]??'')?.[2];
   for(const run of wordRuns(xml,linkFor)){
    // Deleted runs keep their text in w:delText.
    const value=[...run.xml.matchAll(/<w:(?:t|delText)(?:\s[^>]*)?>([\s\S]*?)<\/w:(?:t|delText)>|<w:(tab|br|cr)\b[^>]*\/>/g)].map(item=>item[2]?(item[2]==='tab'?'\t':'\n'):decodeWordXml(item[1])).join('').replace(cell?/\u2029/g:/(?!)/g,'\n');
    text.insert(text.length,value,{...runStyleFor(paragraphStyle&&decodeWordXml(paragraphStyle),wordRunProperties(run.xml),!!heading),...runAttributes(run.xml),...(run.href?{href:run.href}:{}),...(run.revision?{[run.revision.kind]:revisionMark(run.revision)}:{})});
   }
  };
  let tables=0,cells=0;
  for(const match of xml.matchAll(/<w:tbl(?:\s[^>]*?)?(?:\s*\/>|>([\s\S]*?)<\/w:tbl>)|<w:p(?:\s[^>]*?)?(?:\s*\/>|>([\s\S]*?)<\/w:p>)/g)){
   if(/^<w:tbl(?=\s|\/|>)/.test(match[0])){
    if(++tables>WORD_TABLE_LIMITS.tables)throw new Error('Word 表格超过 100 个');
    if(/<w:drawing\b/.test(match[1]??''))throw new Error('表格单元格内图片尚未支持，请移到独立段落');
    const rows=readWordTable(match[1]??'');cells+=rows.length*rows[0].length;if(rows.some(row=>row.some(cell=>cell.paragraphs.length>WORD_TABLE_LIMITS.paragraphs)))throw new Error('单个 Word 单元格最多 100 段');if(cells>WORD_TABLE_LIMITS.cells)throw new Error('Word 单元格超过 2 万');
    const table=createWordTable(rows[0].length,rows.length);paragraphs.push([table]);const data=wordTableRows(table);
    rows.forEach((row,r)=>row.forEach((cell,c)=>{
     if(cell.fill||cell.continued)return;const text=data.get(r).get(c);
     // One cell holds one set of paragraph properties; refuse mixed formatting instead of dropping it.
     const shared=JSON.stringify(paragraphAttributes(cell.paragraphs[0]));if(cell.paragraphs.some(p=>JSON.stringify(paragraphAttributes(p))!==shared))throw new Error('同一 Word 单元格内的段落格式（标题、列表、对齐或缩进）须一致');
     cell.paragraphs.forEach((p,i)=>{if(i)text.insert(text.length,WORD_CELL_PARAGRAPH,{});fill(text,p,i===0,true);});
     if(cell.colSpan>1)text.setAttribute('colSpan',cell.colSpan);if(cell.rowSpan>1)text.setAttribute('rowSpan',cell.rowSpan);
    }));
    const layout=wordTableLayout(table);if([...layout.anchors.values()].reduce((n,a)=>n+a.colSpan*a.rowSpan,0)!==rows.length*rows[0].length)throw new Error('Word 合并单元格区域重叠或无效');
   }else{const source=match[2]??'',hasImage=/<w:drawing\b/.test(source);if(hasImage&&/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/.test(source))throw new Error('图片与文字须放在独立段落，请先拆分图文混排段落');if(!hasImage){const text=new Y.Text();paragraphs.push([text]);fill(text,source);}appendImages(source);}
   if(paragraphs.length>MAX_WORD_PARAGRAPHS)throw new Error('Word 段落超过 10 万');
  }
  const update=Y.encodeStateAsUpdate(doc);validateWordUpdate(null,update);return {doc,update,media:images.media};
 }catch(error){doc.destroy();throw error;}
}
/** Structure of every table that exists in `doc`, keyed by table identity. */
function tableStructures(doc:Y.Doc){return new Map(doc.getArray<WordBlock>('paragraphs').toArray().filter(isWordTable).map(table=>[wordItemIdentity(table),wordTableStructure(table)]));}
/**
 * `lockTables` (team workspaces): tables that existed before may not change rows, columns or merges through a generic
 * update — those go through the table operation endpoint. Creating and deleting whole tables stays allowed.
 */
export function validateWordUpdate(current:Uint8Array|null,update:Uint8Array,options:{lockTables?:boolean;actor?:string}={}){
 const doc=new Y.Doc(),paragraphs=doc.getArray('paragraphs'),settings=doc.getMap('pageSettings');
 try{
  if(current)Y.applyUpdate(doc,current);Y.applyUpdate(doc,update);
  if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('Word 更新缺少依赖，请重新同步');
  validateOfficeComments(doc);if([...doc.share.keys()].some(key=>!['officeComments','officeCommentOperations','paragraphs','pageSettings','wordTableOperations'].includes(key))||paragraphs._map.size)throw new Error('Word 更新包含非法根结构');
  if(doc.share.has('wordTableOperations')){const receipts=doc.getMap('wordTableOperations');if(receipts.size>10000||[...receipts.values()].some(value=>typeof value!=='string'||!/^[a-f0-9]{64}:(\d+-\d+)?$/.test(value)))throw new Error('Word 表格操作回执无效');}
  if(current){const previous=new Y.Doc();try{Y.applyUpdate(previous,current);
   const receipts=(value:Y.Doc)=>JSON.stringify(value.share.has('wordTableOperations')?value.getMap('wordTableOperations').toJSON():{});if(receipts(previous)!==receipts(doc))throw new Error('请通过表格操作接口修改表格结构');
   if(options.lockTables){const after=tableStructures(doc);for(const [table,structure]of tableStructures(previous)){const next=after.get(table);if(next!==undefined&&next!==structure)throw new Error('团队文档的表格行列与合并请通过表格工具修改，请刷新后重试');}}
  }finally{previous.destroy();}}
  // pageSettings: `layout` (paper and margins) and optional `headerFooter`; either may be absent.
  if(settings._start||[...settings.keys()].some(key=>!['layout','headerFooter','trackChanges'].includes(key))||settings.has('trackChanges')&&typeof settings.get('trackChanges')!=='boolean'||settings.has('layout')&&!validWordPage(settings.get('layout'))||settings.has('headerFooter')&&!validWordHeaderFooter(settings.get('headerFooter')))throw new Error('Word 页面设置无效');
  if(paragraphs.length>MAX_WORD_PARAGRAPHS)throw new Error('Word 段落超过 10 万');
  let length=0,fragments=0;const lists=new Map<string,WordList>(),links=new Set<string>(),revisions=new Map<string,string>();
  const texts:Y.Text[]=[],cellTexts=new Set<Y.Text>();let tables=0,cells=0,images=0;
  for(const block of paragraphs.toArray()){
   if(block instanceof Y.Text){texts.push(block);continue;}
   if(isWordImage(block)){validateWordImage(block);if(++images>OFFICE_IMAGE_COUNT)throw new Error('Word 图片超过 1000 个');continue;}
   if(!(block instanceof Y.Map)||++tables>WORD_TABLE_LIMITS.tables)throw new Error('Word 表格结构或数量无效');
   const contents=validateWordTable(block);cells+=contents.length;if(cells>WORD_TABLE_LIMITS.cells)throw new Error('Word 单元格超过 2 万');texts.push(...contents);for(const cell of contents)cellTexts.add(cell);
  }
  if(texts.length>MAX_WORD_PARAGRAPHS)throw new Error('Word 段落超过 10 万');
  for(const value of texts){
   if(!(value instanceof Y.Text))throw new Error('Word 段落结构无效');
   if(Object.entries(value.getAttributes()).some(([key,attribute])=>key==='pageBreakBefore'?typeof attribute!=='boolean':key==='align'?!validWordAlignment(attribute):key==='leftIndent'?!validWordLeftIndent(attribute):key==='rightIndent'?!validWordRightIndent(attribute):key==='firstLineIndent'?!validWordFirstLineIndent(attribute):key==='list'?!validWordList(attribute):key==='lineSpacing'?!validWordLineSpacing(attribute):key==='spaceBefore'||key==='spaceAfter'?!validWordParagraphSpace(attribute):key==='colSpan'||key==='rowSpan'?!cellTexts.has(value)||!Number.isInteger(attribute)||Number(attribute)<2||Number(attribute)>(key==='colSpan'?WORD_TABLE_LIMITS.columns:WORD_TABLE_LIMITS.rows):key!=='heading'||!Number.isInteger(attribute)||Number(attribute)<0||Number(attribute)>3))throw new Error('Word 段落属性无效');
   const list=value.getAttribute('list') as WordList|undefined;if(list){const prior=lists.get(list.id);if(prior&&(prior.kind!==list.kind||prior.start!==list.start))throw new Error('Word 列表定义不一致');lists.set(list.id,list);if(lists.size>10000)throw new Error('Word 列表定义过多');}
   length+=value.length;if(value.length>100_000||length>10_000_000)throw new Error('Word 文本过长');
   if(value.toString().includes(WORD_CELL_PARAGRAPH)&&(!cellTexts.has(value)||value.toString().split(WORD_CELL_PARAGRAPH).length>WORD_TABLE_LIMITS.paragraphs))throw new Error('Word 单元格分段仅可用于表格且每格最多 100 段');
   for(const part of value.toDelta()){
    if(++fragments>200_000)throw new Error('Word 格式片段过多');
    if(typeof part.attributes?.href==='string'){links.add(part.attributes.href);if(links.size>9998)throw new Error('Word 超链接数量过多');}
    if(typeof part.insert!=='string')throw new Error('Word 不支持嵌入对象');
    if(Object.entries(part.attributes??{}).some(([key,attribute])=>key==='fontSize'?!validWordFontSize(attribute):key==='color'?!validWordColor(attribute):key==='highlight'?!validWordHighlight(attribute):key==='vertAlign'?!validWordVertAlign(attribute):key==='fontFamily'?!validWordFontFamily(attribute):key==='href'?!validWordLink(attribute):key==='ins'||key==='del'?!validWordRevisionMark(attribute):!MARKS.includes(key)||typeof attribute!=='boolean'))throw new Error('Word 不支持该富文本属性');
    for(const kind of ['ins','del'] as const){const mark=part.attributes?.[kind];if(validWordRevisionMark(mark)){const known=revisions.get(mark.id);if(known&&known!==mark.author+'\n'+mark.name+'\n'+mark.date)throw new Error('Word 修订标记不一致');revisions.set(mark.id,mark.author+'\n'+mark.name+'\n'+mark.date);}}
   }
  }
  // Revision authorship: a revision id introduced by this update must belong to the requesting member.
  if(options.actor&&current){const previous=new Y.Doc();try{Y.applyUpdate(previous,current);const known=new Set<string>();for(const text of wordBlockTexts(previous.getArray<WordBlock>('paragraphs').toArray()))for(const part of text.toDelta() as {attributes?:Record<string,unknown>}[])for(const kind of WORD_REVISION_KEYS){const mark=part.attributes?.[kind];if(validWordRevisionMark(mark))known.add(mark.id);}
   for(const [id,value]of revisions)if(!known.has(id)&&value.split('\n')[0]!==options.actor)throw new Error('修订标记的作者与当前成员不符');}finally{previous.destroy();}}
 }finally{doc.destroy();}
}
