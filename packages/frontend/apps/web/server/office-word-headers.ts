import {posix} from 'node:path';
import {readOfficeXml} from './office-package';
import {decodeWordXml} from './office-word-link';
import {packageImageReader,type OfficeMedia} from './office-media';
import {validWordHeaderFooter,WORD_PAGE_TOKEN,WORD_PAGES_TOKEN,WORD_HEADER_LINES,WORD_HEADER_SLOTS,WORD_HEADER_IMAGE_MAX,wordHeaderParts,wordHeaderImageKey,emptyWordLine,type WordHeaderFooter,type WordHeaderLine,type WordHeaderPair,type WordHeaderPart,type WordHeaderScope,type WordHeaderSlot} from '../../../../common/nbstore/src/office-word-header';
import type {WordPageLayout} from '../../../../common/nbstore/src/office-word-page';

/** A header/footer picture as stored in `pageSettings.headerImages` (the Word image map fields). */
export type WordHeaderImage={assetId:string;width:number;height:number;alt:string};
const attr=(xml:string,name:string)=>new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];
const fieldToken=(instr:string)=>{const name=instr.trim().split(/\s+/)[0]?.toUpperCase();return name==='PAGE'?WORD_PAGE_TOKEN:name==='NUMPAGES'||name==='SECTIONPAGES'?WORD_PAGES_TOKEN:undefined;};
const clean=(text:string)=>text.replace(new RegExp('[\\x00-\\x09\\x0b-\\x1f\\u2028\\u2029]','g'),' ');
/** Lines of one slot: blank lines dropped, at most five (the rest joined onto the last), at most 255 characters. */
function slotText(lines:string[]){
 const kept=lines.map(line=>clean(line).trim()).filter(Boolean);if(kept.length>WORD_HEADER_LINES)kept.splice(WORD_HEADER_LINES-1,Infinity,kept.slice(WORD_HEADER_LINES-1).join(' '));
 let text=kept.join('\n');if(text.length>255)text=text.slice(0,255).replace(/\n+$/,'');return text;
}
/** Size in pixels from a drawing's extent, scaled down (keeping the ratio) to the header picture limit. */
function headerImageSize(drawing:string){
 const extent=/<wp:extent\b([^>]*?)\/?>/.exec(drawing)?.[1]??'',width=Number(attr(extent,'cx'))/9525,height=Number(attr(extent,'cy'))/9525;
 if(!(width>0&&height>0))return;const scale=Math.min(1,WORD_HEADER_IMAGE_MAX.width/width,WORD_HEADER_IMAGE_MAX.height/height);
 return {width:Math.max(1,Math.round(width*scale)),height:Math.max(1,Math.round(height*scale))};
}
/**
 * One header/footer part as left/centre/right text, one paragraph per line. Tabs split the slots (Word's Header
 * style has a centre and a right tab); line breaks stay inside a slot; PAGE and NUMPAGES fields, simple or complex,
 * become {页码}/{总页数}. The first picture of each slot is kept; tables, shapes and other content are not carried over.
 */
function readLine(xml:string,image:(drawing:string)=>WordHeaderImage|undefined){
 const slots={left:[] as string[],center:[] as string[],right:[] as string[]},images:Partial<Record<WordHeaderSlot,WordHeaderImage>>={};
 // Alternate content repeats a drawing as a fallback; only the first choice is read.
 const source=xml.replace(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g,block=>/<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>/.exec(block)?.[1]??'');
 for(const paragraph of source.matchAll(/<w:p(?:\s[^>]*?)?(?:\/>|>([\s\S]*?)<\/w:p>)/g)){
  const body=paragraph[1]??'',align=attr(/<w:jc\b[^>]*>/.exec(body)?.[0]??'','w:val'),content=body.replace(/<w:pPr\b[\s\S]*?<\/w:pPr>/g,'').replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g,'');
  const segments=[''],pictures:{segment:number;drawing:string}[]=[];let state:'none'|'instr'|'result'='none',instr='',result='';
  const add=(text:string)=>{if(state==='result')result+=text;else if(state==='none')segments[segments.length-1]+=text;};
  for(const token of content.matchAll(/<w:drawing\b[\s\S]*?<\/w:drawing>|<w:fldSimple\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:fldSimple>)|<w:fldChar\b([^>]*?)\/?>|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:tab\b[^>]*\/>|<w:(?:br|cr)\b[^>]*\/>|<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)){
   const [whole,simpleAttrs,simpleBody,charAttrs,instrText,text]=token;
   if(whole.startsWith('<w:drawing')){if(state==='none')pictures.push({segment:segments.length-1,drawing:whole});}
   else if(simpleAttrs!==undefined){const value=fieldToken(decodeWordXml(attr(simpleAttrs,'w:instr')??''));add(value??[...(simpleBody??'').matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(m=>decodeWordXml(m[1])).join(''));}
   else if(charAttrs!==undefined){const type=attr(charAttrs,'w:fldCharType');if(type==='begin'){state='instr';instr='';result='';}else if(type==='separate')state='result';else if(type==='end'){state='none';add(fieldToken(instr)??result);}}
   else if(instrText!==undefined){if(state==='instr')instr+=decodeWordXml(instrText);}
   else if(whole.startsWith('<w:tab')){if(state==='none')segments.push('');}
   else if(whole.startsWith('<w:br')||whole.startsWith('<w:cr')){if(attr(whole,'w:type')!=='page')add('\n');}
   else if(text!==undefined)add(decodeWordXml(text));
  }
  // One segment follows the paragraph alignment; two or more are left/centre/right split by tabs.
  const slotOf=(segment:number,count:number):WordHeaderSlot=>count===1?(align==='center'?'center':align==='right'||align==='end'?'right':'left'):segment===0?'left':segment===1?'center':'right';
  // Pictures are only read once chosen, so skipped ones never enter the document's media.
  for(const {segment,drawing}of pictures){const slot=slotOf(segment,segments.length);if(!images[slot]){const value=image(drawing);if(value)images[slot]=value;}}
  const parts=segments.map(part=>part.trim());if(!parts.some(Boolean))continue;
  if(parts.length===1)slots[slotOf(0,1)].push(parts[0]);
  else{slots.left.push(parts[0]);slots.center.push(parts[1]);if(parts.length>2)slots.right.push(parts.slice(2).filter(Boolean).join(' '));}
 }
 const line={left:slotText(slots.left.flatMap(text=>text.split('\n'))),center:slotText(slots.center.flatMap(text=>text.split('\n'))),right:slotText(slots.right.flatMap(text=>text.split('\n')))};
 return {line,images};
}
/**
 * Default, first-page and even-page headers/footers of the single section, with their pictures (keyed by
 * `wordHeaderImageKey`) and the picture media; undefined when the document has none. Even pages only count when
 * settings.xml turns on evenAndOddHeaders, as in Word. Pictures that cannot be read are skipped.
 */
export function readWordHeaderFooter(files:Record<string,Uint8Array>,documentXml:string,registry?:Map<string,OfficeMedia>):{value?:WordHeaderFooter;images:Record<string,WordHeaderImage>;media:OfficeMedia[]}|undefined{
 const sections=[...documentXml.replace(/<w:sectPrChange\b[^>]*>[\s\S]*?<\/w:sectPrChange>/g,'').matchAll(/<w:sectPr(?=\s|\/|>)(?:[^>]*\/\s*>|[^>]*>([\s\S]*?)<\/w:sectPr>)/g)],section=sections.at(-1)?.[1];if(!section)return;
 const relations=new Map<string,string>(),rels=files['word/_rels/document.xml.rels'];
 if(rels)for(const item of readOfficeXml(rels).matchAll(/<Relationship\b([^>]*)\/?>/g)){const id=attr(item[1],'Id'),target=attr(item[1],'Target');if(id&&target&&attr(item[1],'TargetMode')!=='External'&&!/[\\:#?]/.test(target)){const path=target.startsWith('/')?target.slice(1):posix.normalize(posix.join('word',target));if(path.startsWith('word/')&&!path.split('/').includes('..'))relations.set(id,path);}}
 const media:OfficeMedia[]=[],images:Record<string,WordHeaderImage>={},shared=registry??new Map<string,OfficeMedia>();
 const part=(scope:WordHeaderScope,kind:WordHeaderPart):WordHeaderLine=>{
  const type=scope==='main'?'default':scope;
  for(const ref of section.matchAll(new RegExp(`<w:${kind}Reference\\b([^>]*)\\/?>`,'g'))){
   if((attr(ref[1],'w:type')??'default')!==type)continue;const path=relations.get(attr(ref[1],'r:id')??''),bytes=path&&files[path];if(!bytes)continue;
   const reader=packageImageReader(files,path,shared);
   const image=(drawing:string)=>{const embed=attr(/<a:blip\b[^>]*>/.exec(drawing)?.[0]??'','r:embed'),size=headerImageSize(drawing);if(!embed||!size)return;try{const alt=decodeWordXml(attr(/<wp:docPr\b[^>]*>/.exec(drawing)?.[0]??'','descr')??'').slice(0,1000)||'页眉图片';return {assetId:reader.read(embed),...size,alt};}catch{return;}};
   const {line,images:found}=readLine(readOfficeXml(bytes),image);media.push(...reader.media);
   for(const [slot,value]of Object.entries(found))images[wordHeaderImageKey(scope,kind,slot as WordHeaderSlot)]=value;
   return line;
  }
  return {left:'',center:'',right:''};
 };
 const pair=(scope:WordHeaderScope):WordHeaderPair=>({header:part(scope,'header'),footer:part(scope,'footer')});
 const flag=(xml:string,name:string)=>{const match=new RegExp(`<w:${name}(?=[\\s/>])([^>]*)\\/?>`).exec(xml);return !!match&&!['0','false','off'].includes(attr(match[1],'w:val')??'1');};
 const settings=files['word/settings.xml']?readOfficeXml(files['word/settings.xml']):'',firstPage=flag(section,'titlePg'),evenPages=flag(settings,'evenAndOddHeaders');
 const value:WordHeaderFooter={...pair('main'),...(firstPage?{first:pair('first')}:{}),...(evenPages?{even:pair('even')}:{})};
 if(!firstPage&&!evenPages&&emptyWordLine(value.header)&&emptyWordLine(value.footer)&&!Object.keys(images).length)return;
 // Media stays with the result even if the text were invalid: the body may reuse a picture first read here.
 return validWordHeaderFooter(value)?{value,images,media}:{images:{},media};
}

const xmlText=(value:string)=>value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
/** An inline picture run (DrawingML); `id` must be unique among the document's drawings. */
export function wordInlineImageRun(relationId:string,fileName:string,id:number,image:{width:number;height:number;alt:string}){
 const cx=Math.round(image.width*9525),cy=Math.round(image.height*9525),alt=xmlText(image.alt);
 return `<w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="图片 ${id}" descr="${alt}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${fileName}" descr="${alt}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
export const WORD_DRAWING_NAMESPACES='xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
type ExportImage={image:WordHeaderImage;file:string};
/**
 * One header/footer part: one paragraph per line with centre and right tab stops at the middle and right edge of the
 * text area. A slot's picture is its first line; slots with fewer lines leave that tab position empty.
 */
function partXml(kind:'hdr'|'ftr',line:WordHeaderLine,images:Partial<Record<WordHeaderSlot,ExportImage>>,page:WordPageLayout,drawingId:()=>number){
 const width=page.width-page.left-page.right,relations:string[]=[],runs=(text:string)=>wordHeaderParts(text).map(part=>'field' in part?`<w:fldSimple w:instr=" ${part.field} "><w:r><w:t>1</w:t></w:r></w:fldSimple>`:`<w:r><w:t xml:space="preserve">${xmlText(part.text)}</w:t></w:r>`).join('');
 const cells=Object.fromEntries(WORD_HEADER_SLOTS.map(slot=>{
  const entry=images[slot],items=line[slot]?line[slot].split('\n').map(runs):[];
  if(entry){const id='rIdImg'+(relations.length+1);relations.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${entry.file}"/>`);items.unshift(wordInlineImageRun(id,entry.file,drawingId(),entry.image));}
  return [slot,items];
 })) as Record<WordHeaderSlot,string[]>;
 const rows=Math.max(1,...WORD_HEADER_SLOTS.map(slot=>cells[slot].length)),tab='<w:r><w:tab/></w:r>';
 const paragraphs=Array.from({length:rows},(_,row)=>{const [left,center,right]=WORD_HEADER_SLOTS.map(slot=>cells[slot][row]??'');return `<w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(width/2)}"/><w:tab w:val="right" w:pos="${width}"/></w:tabs></w:pPr>${left+(center||right?tab+center:'')+(right?tab+right:'')}</w:p>`;}).join('');
 return {xml:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ${WORD_DRAWING_NAMESPACES}>${paragraphs}</w:${kind}>`,relations};
}
/**
 * Parts, relationships, content types and sectPr references for export; empty when there is nothing to write.
 * `image` returns the media file name (under word/media/) of a stored picture and `drawingId` the next drawing id.
 */
export function wordHeaderFooterParts(value:WordHeaderFooter|undefined,page:WordPageLayout,pictures:Record<string,WordHeaderImage>={},image:(assetId:string)=>string=()=>{throw new Error('页眉图片缺失');},drawingId:()=>number=()=>1){
 const files:Record<string,string>={},relations:string[]=[],types:string[]=[],references:string[]=[];if(!value)return {files,relations,types,references,titlePage:false,evenPages:false};
 const add=(kind:WordHeaderPart,scope:WordHeaderScope,line:WordHeaderLine,n:number)=>{
  const slots=Object.fromEntries(WORD_HEADER_SLOTS.flatMap(slot=>{const picture=pictures[wordHeaderImageKey(scope,kind,slot)];return picture?[[slot,{image:picture,file:image(picture.assetId)}]]:[];}));
  const name=`${kind}${n}.xml`,id=`rId${kind==='header'?'Hdr':'Ftr'}${n}`,part=partXml(kind==='header'?'hdr':'ftr',line,slots,page,drawingId);files['word/'+name]=part.xml;
  if(part.relations.length)files[`word/_rels/${name}.rels`]=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${part.relations.join('')}</Relationships>`;
  relations.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${name}"/>`);
  types.push(`<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`);references.push(`<w:${kind}Reference w:type="${scope==='main'?'default':scope}" r:id="${id}"/>`);
 };
 const has=(scope:WordHeaderScope,kind:WordHeaderPart)=>WORD_HEADER_SLOTS.some(slot=>pictures[wordHeaderImageKey(scope,kind,slot)]);
 // First-page and even-page variants are written even when empty, so a blank first or even page stays blank.
 for(const kind of ['header','footer'] as const)if(!emptyWordLine(value[kind])||has('main',kind))add(kind,'main',value[kind],1);
 if(value.first){add('header','first',value.first.header,2);add('footer','first',value.first.footer,2);}
 if(value.even){add('header','even',value.even.header,3);add('footer','even',value.even.footer,3);}
 // CT_SectPr: header references before footer references.
 references.sort((a,b)=>Number(b.startsWith('<w:header'))-Number(a.startsWith('<w:header')));
 return {files,relations,types,references,titlePage:!!value.first,evenPages:!!value.even};
}
