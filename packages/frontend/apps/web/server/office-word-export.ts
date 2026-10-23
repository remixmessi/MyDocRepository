import {unpackOfficeMedia,checkOfficeMedia} from './office-media';
import {isWordImage,wordHeaderImages} from '../../../../common/nbstore/src/office-image';
import {wordBlockTexts,wordTableRows,wordTableLayout,WORD_CELL_PARAGRAPH,type WordBlock} from '../../../../common/nbstore/src/office-word-table';
import {DEFAULT_WORD_PAGE,type WordPageLayout} from '../../../../common/nbstore/src/office-word-page';
import type {WordList} from '../../../../common/nbstore/src/office-word-list';
import * as Y from 'yjs';
import {strToU8,zipSync} from 'fflate';
import {validateWordUpdate} from './office-word';
import {validOfficeComment,type OfficeCommentThread} from '../../../../common/nbstore/src/office-comments';
import {officeAnchorTarget} from '../../../../common/nbstore/src/office-comment-anchor';
import {wordHeaderFooterParts,wordInlineImageRun} from './office-word-headers';
import {validWordHeaderFooter} from '../../../../common/nbstore/src/office-word-header';
import {validWordRevisionMark,validWordRunFormatMark,validWordParagraphFormatMark,type WordRevisionMark} from '../../../../common/nbstore/src/office-word-revision';

type CommentMark={offset:number;kind:'start'|'end';id:number};
/**
 * File comments become native Word comments: each thread is one comment, replies are threaded comments
 * (commentsExtended paraIdParent), and resolved threads are marked done. Anchors that no longer resolve, and
 * threads without a text anchor, attach to the start of the document so no discussion is lost.
 */
function wordComments(doc:Y.Doc,texts:Y.Text[]){
 const threads=doc.share.has('officeComments')?[...doc.getMap<OfficeCommentThread>('officeComments').values()].filter(validOfficeComment).sort((a,b)=>a.messages[0].created-b.messages[0].created||a.id.localeCompare(b.id)):[];
 const marks=new Map<Y.Text,CommentMark[]>(),add=(text:Y.Text,mark:CommentMark)=>{const list=marks.get(text)??[];list.push(mark);marks.set(text,list);};
 const comments:string[]=[],extended:string[]=[],floating:number[]=[];let id=0,para=0;
 const paraId=()=>(0x10000000+ ++para).toString(16).toUpperCase();
 for(const thread of threads){
  const target=thread.anchor?officeAnchorTarget(doc,thread.anchor):undefined,range=target?.kind==='text'&&texts.includes(target.first)&&texts.includes(target.last)?target:undefined;
  let parent='';
  for(const message of thread.messages){
   const current=id++,pid=paraId(),initials=initialsOf(message.authorName),date=new Date(message.created).toISOString().replace(/\.\d{3}Z$/,'Z');
   const body=message.body.split('\n').map((line,i,lines)=>`<w:p${i===lines.length-1?` w14:paraId="${pid}"`:''}>${i===0?'<w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:annotationRef/></w:r>':''}<w:r><w:t xml:space="preserve">${xmlText(line)}</w:t></w:r></w:p>`).join('');
   comments.push(`<w:comment w:id="${current}" w:author="${xmlText(message.authorName)}" w:date="${date}" w:initials="${xmlText(initials)}">${body}</w:comment>`);
   extended.push(`<w15:commentEx w15:paraId="${pid}"${parent?` w15:paraIdParent="${parent}"`:''} w15:done="${thread.resolved?1:0}"/>`);
   if(!parent)parent=pid;
   // Replies share the root comment's range so Word shows them in the same thread.
   if(range){add(range.first,{offset:range.start,kind:'start',id:current});add(range.last,{offset:range.end,kind:'end',id:current});}else floating.push(current);
  }
 }
 for(const list of marks.values())list.sort((a,b)=>a.offset-b.offset||(a.kind===b.kind?a.id-b.id:a.kind==='end'?-1:1));
 return {marks,floating,comments,extended,count:id};
}
/** "Lisa Chen" → "LC", "王小明" → "王" (Word's own convention for CJK names). */
const initialsOf=(name:string)=>{const words=name.trim().split(/\s+/).filter(Boolean);if(!words.length)return '?';return /^[\x00-\x7f]/.test(words[0])?words.slice(0,3).map(word=>[...word][0].toUpperCase()).join(''):[...words[0]][0];};
const commentStart=(id:number)=>`<w:commentRangeStart w:id="${id}"/>`;
const commentEnd=(id:number)=>`<w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;

function xmlText(value:string){
 for(const character of value){const code=character.codePointAt(0)!;if(!(code===9||code===10||code===13||code>=32&&code<=0xd7ff||code>=0xe000&&code<=0xfffd||code>=0x10000&&code<=0x10ffff))throw new Error('正文包含无法写入 DOCX 的字符');}
 return value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&apos;').replaceAll('\r','&#13;');
}
/** Exports only the declared paragraph model and supported direct character marks, never alters the original. */
export function exportDocxWord(snapshot:Uint8Array){
 const payload=unpackOfficeMedia(snapshot);snapshot=payload.state;const media=checkOfficeMedia(snapshot,'word',payload.media),imageFiles:Record<string,Uint8Array>={},imageRelations:string[]=[];let drawingId=0;
 validateWordUpdate(null,snapshot);
 const doc=new Y.Doc();
 try{
  Y.applyUpdate(doc,snapshot);
  const page=(doc.getMap('pageSettings').get('layout')??DEFAULT_WORD_PAGE) as WordPageLayout;
  const blocks=doc.getArray<WordBlock>('paragraphs').toArray(),paragraphs=wordBlockTexts(blocks);
  const lists=new Map<string,{list:WordList;number:number}>();
  // Lists named only by a formatting revision's previous value still need a numbering definition.
  for(const paragraph of paragraphs){const format=paragraph.getAttribute('pfmt'),previous=validWordParagraphFormatMark(format)?format.old.list:undefined;for(const list of [paragraph.getAttribute('list'),previous] as (WordList|undefined|null)[])if(list&&!lists.has(list.id))lists.set(list.id,{list,number:lists.size+1});}
  const links=new Map<string,string>(),notes=wordComments(doc,paragraphs),firstText=paragraphs[0];let revisionId=notes.count+1;
  const trackChanges=doc.getMap('pageSettings').get('trackChanges')===true;
  // Pictures are stored once under word/media/ even when the body and several headers use them; each part relates to the file itself.
  const mediaNames=new Map<string,string>(),bodyRelations=new Map<string,string>();
  const mediaFile=(assetId:string)=>{const asset=media.get(assetId);if(!asset)throw new Error('Word 图片缺失');let name=mediaNames.get(assetId);if(!name){name='image'+(mediaNames.size+1)+(asset.entry.mime==='image/png'?'.png':'.jpg');mediaNames.set(assetId,name);imageFiles['word/media/'+name]=asset.bytes;}return name;};
  const bodyImage=(name:string)=>{let id=bodyRelations.get(name);if(!id){id='rImg'+(bodyRelations.size+1);bodyRelations.set(name,id);imageRelations.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);}return id;};
  const revisionAttributes=(mark:WordRevisionMark)=>`w:id="${revisionId++}" w:author="${xmlText(mark.name)}" w:date="${new Date(mark.date).toISOString().replace(/\.\d{3}Z$/,'Z')}"`;
  /** Current values with a revision's recorded previous values put back (null = was not set). */
  const previousValues=(current:Record<string,unknown>,old:Record<string,unknown>)=>{const result={...current};for(const [key,value]of Object.entries(old)){if(value===null)delete result[key];else result[key]=value;}return result;};
  const runProperties=(marks:Record<string,unknown>)=>{
   // CT_RPr element order: rFonts, b, i, strike, color, sz, highlight, u, vertAlign.
   const fonts=typeof marks.fontFamily==='string'?`<w:rFonts w:ascii="${xmlText(marks.fontFamily)}" w:hAnsi="${xmlText(marks.fontFamily)}" w:eastAsia="${xmlText(marks.fontFamily)}"/>`:'';
   const toggle=[['bold','b'],['italic','i'],['strike','strike']].filter(([key])=>marks[key]!==undefined).map(([key,tag])=>`<w:${tag} w:val="${marks[key]?'1':'0'}"/>`).join('');
   const color=typeof marks.color==='string'?`<w:color w:val="${marks.color==='auto'?'auto':marks.color.slice(1).toUpperCase()}"/>`:'';
   const size=typeof marks.fontSize==='number'?`<w:sz w:val="${marks.fontSize*2}"/>`:'';
   const highlight=typeof marks.highlight==='string'?`<w:highlight w:val="${marks.highlight}"/>`:'';
   const underline=marks.underline===undefined?'':`<w:u w:val="${marks.underline?'single':'none'}"/>`;
   const vert=typeof marks.vertAlign==='string'?`<w:vertAlign w:val="${marks.vertAlign}"/>`:'';
   return fonts+toggle+color+size+highlight+underline+vert;
  };
  const paragraphProperties=(attributes:Record<string,unknown>)=>{
   const heading=Number(attributes.heading??0),list=attributes.list as WordList|undefined;
   const numbering=list?`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${lists.get(list.id)!.number}"/></w:numPr>`:'';
   const align=attributes.align as string|undefined,leftIndent=attributes.leftIndent as number|undefined;
   const rightIndent=attributes.rightIndent as number|undefined,firstLineIndent=attributes.firstLineIndent as number|undefined;
   const ind=(leftIndent===undefined?'':` w:left="${leftIndent}"`)+(rightIndent===undefined?'':` w:right="${rightIndent}"`)+(firstLineIndent===undefined?'':firstLineIndent<0?` w:hanging="${-firstLineIndent}"`:` w:firstLine="${firstLineIndent}"`);
   const lineSpacing=attributes.lineSpacing as number|undefined,spaceBefore=attributes.spaceBefore as number|undefined,spaceAfter=attributes.spaceAfter as number|undefined;
   const spacing=(spaceBefore===undefined?'':` w:before="${spaceBefore}"`)+(spaceAfter===undefined?'':` w:after="${spaceAfter}"`)+(lineSpacing===undefined?'':` w:line="${Math.round(lineSpacing*240)}" w:lineRule="auto"`);
   // CT_PPr order: spacing precedes ind and jc.
   const layout=(spacing?`<w:spacing${spacing}/>`:'')+(ind?`<w:ind${ind}/>`:'')+(align?`<w:jc w:val="${align==='justify'?'both':align}"/>`:'');
   return `<w:pStyle w:val="${heading?'Heading'+heading:'Normal'}"/>${attributes.pageBreakBefore===undefined?'':`<w:pageBreakBefore w:val="${attributes.pageBreakBefore?1:0}"/>`}${numbering}${layout}${heading?`<w:outlineLvl w:val="${heading-1}"/>`:''}`;
  };
  // A paragraph's insBreak/delBreak is the paragraph mark before it, which Word stores on the previous paragraph.
  const breakMarks=new Map<Y.Text,{kind:'ins'|'del';mark:WordRevisionMark}>();
  blocks.forEach((block,i)=>{const previous=blocks[i-1];if(!(block instanceof Y.Text)||!(previous instanceof Y.Text))return;for(const [key,kind]of [['insBreak','ins'],['delBreak','del']] as const){const mark=block.getAttribute(key);if(validWordRevisionMark(mark))breakMarks.set(previous,{kind,mark});}});
  const paragraphXml=(paragraph:Y.Text,cell=false)=>{
   // Cells keep their paragraphs in one Y.Text split by U+2029; every paragraph shares the cell's paragraph properties.
   const delta=paragraph.toDelta() as {insert:string;attributes?:Record<string,boolean|string|number>}[],pieces:typeof delta[]=[[]],starts=[0];let offset=0;
   for(const part of delta)String(part.insert).split(cell?WORD_CELL_PARAGRAPH:/(?!)/).forEach((text,i)=>{if(i){offset++;pieces.push([]);starts.push(offset);}if(text)pieces[pieces.length-1].push({...part,insert:text});offset+=text.length;});
   return pieces.map((piece,i)=>paragraphPiece(paragraph,piece,starts[i],i===pieces.length-1)).join('');
  };
  const paragraphPiece=(paragraph:Y.Text,source:{insert:string;attributes?:Record<string,boolean|string|number>}[],base=0,last=true)=>{
   // Split runs where comment ranges begin or end; marks past this piece belong to a later cell paragraph.
   // A mark exactly at the end of a cell paragraph stays with that paragraph.
   const end=base+source.reduce((n,part)=>n+part.insert.length,0),pending=(notes.marks.get(paragraph)??[]).filter(mark=>mark.offset>=base&&(last||mark.offset<=end));
   const delta:({insert:string;attributes?:Record<string,boolean|string|number>}|{mark:CommentMark})[]=[];let at=base,next=0;
   const flush=(limit:number)=>{while(next<pending.length&&pending[next].offset<=limit)delta.push({mark:pending[next++]});};
   // After flush(at) every remaining mark lies beyond `at`, so each cut is at least one character.
   for(const part of source){let text=part.insert;while(text){flush(at);const cut=next<pending.length?Math.min(text.length,pending[next].offset-at):text.length;delta.push({...part,insert:text.slice(0,cut)});text=text.slice(cut);at+=cut;}}
   flush(Number.MAX_SAFE_INTEGER);
   if(paragraph===firstText&&base===0)delta.unshift(...notes.floating.flatMap(id=>[{mark:{offset:0,kind:'start' as const,id}},{mark:{offset:0,kind:'end' as const,id}}]));
   const runs=delta.map(part=>{
    if('mark' in part)return part.mark.kind==='start'?commentStart(part.mark.id):commentEnd(part.mark.id);
    const marks=part.attributes??{};
    // A formatting revision writes the previous properties in w:rPrChange, last inside w:rPr.
    const format=validWordRunFormatMark(marks.fmt)?marks.fmt:undefined;
    const properties=runProperties(marks)+(format?`<w:rPrChange ${revisionAttributes(format)}><w:rPr>${runProperties(previousValues(marks,format.old))}</w:rPr></w:rPrChange>`:'');
    // Deleted text is written as w:delText inside w:del; insertions wrap the run in w:ins (inside any hyperlink).
    const deleted=validWordRevisionMark(marks.del),textTag=deleted?'w:delText':'w:t';
    const text=(part.insert as string).split(/([\t\n])/).map(value=>value==='\t'?'<w:tab/>':value==='\n'?'<w:br/>':`<${textTag} xml:space="preserve">${xmlText(value)}</${textTag}>`).join('');
    let run=`<w:r>${properties?`<w:rPr>${properties}</w:rPr>`:''}${text}</w:r>`;
    for(const kind of ['ins','del'] as const){const mark=marks[kind];if(!validWordRevisionMark(mark))continue;run=`<w:${kind} ${revisionAttributes(mark)}>${run}</w:${kind}>`;}
    if(typeof marks.href!=='string')return run;
    let id=links.get(marks.href);if(!id){if(links.size>=9998)throw new Error('Word 超链接数量过多');id='rIdLink'+(links.size+1);links.set(marks.href,id);}
    return `<w:hyperlink r:id="${id}">${run}</w:hyperlink>`;
   }).join('');
   // CT_PPr: properties, then the paragraph mark's w:rPr (its own insertion/deletion), then w:pPrChange.
   const attributes=paragraph.getAttributes(),format=validWordParagraphFormatMark(attributes.pfmt)?attributes.pfmt:undefined,mark=last?breakMarks.get(paragraph):undefined;
   const properties=paragraphProperties(attributes)+(mark?`<w:rPr><w:${mark.kind} ${revisionAttributes(mark.mark)}/></w:rPr>`:'')+(format?`<w:pPrChange ${revisionAttributes(format)}><w:pPr>${paragraphProperties(previousValues(attributes,format.old))}</w:pPr></w:pPrChange>`:'');
   return `<w:p><w:pPr>${properties}</w:pPr>${runs}</w:p>`;
  };
  const content=blocks.map(block=>{
   if(block instanceof Y.Text)return paragraphXml(block);
   if(block.get('kind')==='image'){const name=mediaFile(String(block.get('assetId'))),id=bodyImage(name);let run=wordInlineImageRun(id,name,++drawingId,{width:Number(block.get('width')),height:Number(block.get('height')),alt:String(block.get('alt'))});for(const kind of ['ins','del'] as const){const mark=block.get(kind);if(validWordRevisionMark(mark))run=`<w:${kind} ${revisionAttributes(mark)}>${run}</w:${kind}>`;}return `<w:p>${run}</w:p>`;}

   const rows=wordTableRows(block),columns=Number(block.get('columns')),layout=wordTableLayout(block);if(!rows.length)throw new Error('空表格无法导出，请先添加一行或删除表格');
   const available=page.width-page.left-page.right,width=Math.floor(available/columns);
   const widths=Array.from({length:columns},(_,i)=>i===columns-1?available-width*(columns-1):width);
   const borders=['top','left','bottom','right','insideH','insideV'].map(edge=>`<w:${edge} w:val="single" w:sz="4" w:color="auto"/>`).join('');
   return `<w:tbl><w:tblPr><w:tblW w:w="${available}" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.toArray().map((row,r)=>`<w:tr>${row.toArray().map((_,c)=>{const owner=layout.owner[r][c]!,span=layout.anchors.get(owner)!;if(span.column!==c)return '';const width=widths.slice(c,c+span.colSpan).reduce((a,b)=>a+b,0),grid=span.colSpan>1?`<w:gridSpan w:val="${span.colSpan}"/>`:'',merge=span.rowSpan>1||span.row!==r?`<w:vMerge${span.row===r?' w:val="restart"':''}/>`:'';return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${grid}${merge}</w:tcPr>${span.row===r?paragraphXml(owner,true):'<w:p/>'}</w:tc>`;}).join('')}</w:tr>`).join('')}</w:tbl>`;
  }).join('');
  // Headers are built after the body so their drawing ids follow the body's.
  const headerFooter=doc.getMap('pageSettings').get('headerFooter'),pictures=Object.fromEntries([...wordHeaderImages(doc)?.entries()??[]].map(([key,image])=>[key,{assetId:String(image.get('assetId')),width:Number(image.get('width')),height:Number(image.get('height')),alt:String(image.get('alt'))}]));
  const headers=wordHeaderFooterParts(validWordHeaderFooter(headerFooter)?headerFooter:undefined,page,pictures,mediaFile,()=>++drawingId);
  // Without any text paragraph (empty, or images only) unanchored comments get their own leading paragraph.
  const body=(!firstText&&notes.floating.length?`<w:p>${notes.floating.map(id=>commentStart(id)+commentEnd(id)).join('')}</w:p>`:'')+content||'<w:p></w:p>';
  const commentsXml=notes.count?`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${notes.comments.join('')}</w:comments>`:'',extendedXml=notes.count?`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">${notes.extended.join('')}</w15:commentsEx>`:'';
  if(strToU8(commentsXml).length>8*1024*1024)throw new Error('导出批注 XML 超过 8 MB');
  const xml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr>${headers.references.join('')}<w:pgSz w:w="${page.width}" w:h="${page.height}" w:orient="${page.width>page.height?'landscape':'portrait'}"/><w:pgMar w:top="${page.top}" w:right="${page.right}" w:bottom="${page.bottom}" w:left="${page.left}" w:header="720" w:footer="720" w:gutter="0"/>${headers.titlePage?'<w:titlePg/>':''}</w:sectPr></w:body></w:document>`;
  const document=strToU8(xml);if(document.length>8*1024*1024)throw new Error('导出正文 XML 超过 8 MB');
  const numbering=lists.size?`<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${[...lists.values()].map(({list,number})=>`<w:abstractNum w:abstractNumId="${number}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="${list.start}"/><w:numFmt w:val="${list.kind}"/><w:lvlText w:val="${list.kind==='bullet'?'•':'%1.'}"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`).join('')}${[...lists.values()].map(({number})=>`<w:num w:numId="${number}"><w:abstractNumId w:val="${number}"/></w:num>`).join('')}</w:numbering>`:'';
  if(strToU8(numbering).length>8*1024*1024)throw new Error('导出列表 XML 超过 8 MB');
  // CT_Settings order: trackRevisions precedes evenAndOddHeaders.
  const settingsXml=trackChanges||headers.evenPages?`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${trackChanges?'<w:trackRevisions/>':''}${headers.evenPages?'<w:evenAndOddHeaders/>':''}</w:settings>`:'';
  const relationships=strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'+(lists.size?'<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>':'')+headers.relations.join('')+(settingsXml?'<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>':'')+(notes.count?'<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>':'')+[...links].map(([href,id])=>`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlText(href)}" TargetMode="External"/>`).join('')+imageRelations.join('')+'</Relationships>');
  if(relationships.length>8*1024*1024)throw new Error('导出关系 XML 超过 8 MB');
  const result=zipSync({
   ...imageFiles,
   '[Content_Types].xml':strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'+(lists.size?'<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>':'')+headers.types.join('')+(settingsXml?'<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>':'')+(notes.count?'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>':'')+'</Types>'),
   '_rels/.rels':strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
   'word/document.xml':document,
   ...(numbering?{'word/numbering.xml':strToU8(numbering)}:{}),
   ...Object.fromEntries(Object.entries(headers.files).map(([name,xml])=>[name,strToU8(xml)])),
   ...(settingsXml?{'word/settings.xml':strToU8(settingsXml)}:{}),
   ...(notes.count?{'word/comments.xml':strToU8(commentsXml),'word/commentsExtended.xml':strToU8(extendedXml)}:{}),
   'word/_rels/document.xml.rels':relationships,
   'word/styles.xml':strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>${[1,2,3].map(level=>`<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="${level-1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${[32,28,24][level-1]}"/></w:rPr></w:style>`).join('')}</w:styles>`),
  },{level:6});
  return result;
 }finally{doc.destroy();}
}
