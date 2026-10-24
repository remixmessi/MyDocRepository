import {posix} from 'node:path';
import {readOfficeXml} from './office-package';
import {decodeWordXml} from './office-word-link';
import {validWordWatermark,wordWatermarkSize,WORD_NOTE_LIMITS,type WordWatermark,type WordTocLevels} from '../../../../common/nbstore/src/office-word-reference';
import type {WordPageLayout} from '../../../../common/nbstore/src/office-word-page';
const attr=(xml:string,name:string)=>new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];
const xmlText=(value:string)=>value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
export const WORD_VML_NAMESPACES='xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

// ---- watermark --------------------------------------------------------------------------------------------------
/** Word's own text watermark: a VML WordArt shape (PowerPlusWaterMarkObject) in the header, centred on the margins. */
export function wordWatermarkXml(watermark:WordWatermark,page:WordPageLayout,id:number){
 const size=wordWatermarkSize(watermark,page),ems=[...watermark.text].reduce((total,char)=>total+(/[\u2E80-\uFFFF]/.test(char)?1:.6),0),width=Math.round(size*Math.max(1,ems)),height=Math.round(size*1.15);
 const style=`position:absolute;margin-left:0;margin-top:0;width:${width}pt;height:${height}pt;${watermark.layout==='diagonal'?'rotation:315;':''}z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin`;
 const shapetype='<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';
 return `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>${shapetype}<v:shape id="PowerPlusWaterMarkObject${id}" o:spid="_x0000_s${2049+id}" type="#_x0000_t136" style="${style}" o:allowincell="f" fillcolor="${watermark.color.toLowerCase()}" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;宋体&quot;;font-size:${size}pt" string="${xmlText(watermark.text)}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r></w:p>`;
}
const NAMED_COLORS:Record<string,string>={silver:'#C0C0C0',gray:'#808080',grey:'#808080',black:'#000000',red:'#FF0000',blue:'#0000FF',green:'#008000',white:'#FFFFFF',yellow:'#FFFF00'};
/** The first text watermark in the given header parts (Word's VML WordArt; picture watermarks are not read). */
export function readWordWatermark(headers:string[],page:WordPageLayout):WordWatermark|undefined{
 for(const xml of headers)for(const shape of xml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)){
  const path=/<v:textpath\b([^>]*)\/?>/.exec(shape[2]);if(!path)continue;const id=attr(shape[1],'id')??'';if(!/WaterMark/i.test(id)&&!/<v:fill\b[^>]*opacity/.test(shape[2]))continue;
  const text=decodeWordXml(attr(path[1],'string')??'').trim(),raw=(attr(shape[1],'fillcolor')??'silver').trim().toLowerCase(),hex=/^#[0-9a-f]{6}$/.test(raw)?raw.toUpperCase():/^#[0-9a-f]{3}$/.test(raw)?'#'+[...raw.slice(1)].map(c=>c+c).join('').toUpperCase():NAMED_COLORS[raw.split(/\s/)[0]]??'#C0C0C0';
  const style=attr(shape[1],'style')??'',rotation=Number(/rotation:\s*(-?\d+(?:\.\d+)?)/.exec(style)?.[1]??0),pathStyle=decodeWordXml(attr(path[1],'style')??''),declared=Number(/font-size:\s*(\d+(?:\.\d+)?)pt/.exec(pathStyle)?.[1]);
  const value:WordWatermark={text:[...text].slice(0,50).join(''),color:hex,layout:Math.abs(rotation%180)>=20&&Math.abs(rotation%180)<=160?'diagonal':'horizontal'};
  // Word writes font-size:1pt and lets the shape size the text; only a real size is kept, otherwise it fits the page.
  if(Number.isFinite(declared)&&declared>=20&&declared<=200)value.size=Math.round(declared);
  if(validWordWatermark(value)){if(value.size&&value.size===wordWatermarkSize({...value,size:undefined},page))delete value.size;return value;}
 }
}

// ---- table of contents ------------------------------------------------------------------------------------------
export type WordTocEntry={level:number;title:string;bookmark:string};
/**
 * A Word TOC field inside a "Table of Contents" content control. Entries are pre-filled (linked to bookmarks on the
 * headings) without page numbers, and settings.xml asks Word to update fields on open, which fills them in.
 */
export function wordTocXml(entries:WordTocEntry[],levels:WordTocLevels,width:number){
 const instr=`TOC \\o "1-${levels}" \\h \\z \\u`,shown=entries.filter(entry=>entry.level<=levels),tab=`<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${width}"/></w:tabs>`;
 const begin=`<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`,end='<w:r><w:fldChar w:fldCharType="end"/></w:r>';
 const rows=shown.length?shown.map((entry,i)=>`<w:p><w:pPr><w:pStyle w:val="TOC${entry.level}"/>${tab}</w:pPr>${i===0?begin:''}<w:hyperlink w:anchor="${entry.bookmark}" w:history="1"><w:r><w:t xml:space="preserve">${xmlText(entry.title)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${entry.bookmark} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>${i===shown.length-1?end:''}</w:p>`).join('')
  :`<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${begin}<w:r><w:t>未找到目录项。</w:t></w:r>${end}</w:p>`;
 return `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>目录</w:t></w:r></w:p>${rows}</w:sdtContent></w:sdt>`;
}
export const WORD_TOC_MARKER='mydoc-toc';
const tocLevels=(instr:string)=>{const match=/\\o\s*"?\s*(\d)\s*-\s*(\d)/.exec(instr);return Math.min(3,Math.max(1,Number(match?.[2]??3))) as WordTocLevels;};
/**
 * Replace each table of contents (a "Table of Contents" content control or a bare TOC field spanning paragraphs) by
 * one marker paragraph, so the body reader turns it into a TOC block instead of importing its entries as text.
 */
export function replaceWordTocs(xml:string){
 const marker=(levels:WordTocLevels)=>`<w:p><w:pPr><w:pStyle w:val="${WORD_TOC_MARKER}-${levels}"/></w:pPr></w:p>`;
 let out=xml.replace(/<w:sdt\b[^>]*>(?:(?!<w:sdt\b)[\s\S])*?<w:docPartGallery\b[^>]*w:val\s*=\s*(["'])Table of Contents\1[\s\S]*?<\/w:sdt>/g,block=>{const instr=/<w:instrText\b[^>]*>([^<]*\bTOC\b[^<]*)<\/w:instrText>/.exec(block)?.[1];return marker(tocLevels(decodeWordXml(instr??'')));});
 for(let guard=0;guard<20;guard++){
  const instr=/<w:instrText\b[^>]*>\s*TOC\b([^<]*)<\/w:instrText>/.exec(out);if(!instr)break;
  const before=out.lastIndexOf('fldCharType="begin"',instr.index),begin=out.lastIndexOf('<w:p',before);if(before<0||begin<0)break;
  let depth=0,end=-1;for(const token of out.slice(before-20).matchAll(/fldCharType\s*=\s*["'](begin|end)["']/g)){depth+=token[1]==='begin'?1:-1;if(!depth){end=before-20+token.index!;break;}}if(end<0)break;
  const close=out.indexOf('</w:p>',end);if(close<0)break;out=out.slice(0,begin)+marker(tocLevels(decodeWordXml(instr[1])))+out.slice(close+6);
 }
 return out;
}

// ---- footnotes and endnotes -------------------------------------------------------------------------------------
/** Footnote or endnote bodies by Word id (separators skipped); paragraphs become lines. */
export function readWordNoteBodies(files:Record<string,Uint8Array>,kind:'footnote'|'endnote'){
 const bodies=new Map<string,string>(),rels=files['word/_rels/document.xml.rels'];let path=`word/${kind}s.xml`;
 if(rels)for(const item of readOfficeXml(rels).matchAll(/<Relationship\b([^>]*)\/?>/g)){if(attr(item[1],'Type')?.endsWith('/'+kind+'s')){const target=attr(item[1],'Target')??'';if(!/[\\:#?]/.test(target)){const resolved=target.startsWith('/')?target.slice(1):posix.normalize(posix.join('word',target));if(resolved.startsWith('word/')&&!resolved.split('/').includes('..'))path=resolved;}}}
 if(!files[path])return bodies;const xml=readOfficeXml(files[path]);
 for(const note of xml.matchAll(new RegExp(`<w:${kind}\\b([^>]*)>([\\s\\S]*?)<\\/w:${kind}>`,'g'))){
  const type=attr(note[1],'w:type'),id=attr(note[1],'w:id');if(!id||type&&type!=='normal')continue;
  const lines=[...note[2].matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>|<w:p\b[^>]*\/>/g)].map(p=>[...(p[1]??'').replace(/<w:(?:footnoteRef|endnoteRef)\b[^>]*\/>/g,'').matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:(tab)\b[^>]*\/>|<w:(?:br|cr)\b[^>]*\/>/g)].map(m=>m[1]!==undefined?decodeWordXml(m[1]):m[2]?'\t':'\n').join(''));
  let text=lines.join('\n').replace(/^\s+/,'').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\u2029\uE000]/g,'');if(text.length>WORD_NOTE_LIMITS.text)text=text.slice(0,WORD_NOTE_LIMITS.text-1)+'…';bodies.set(id,text);
 }
 return bodies;
}
/** footnotes.xml / endnotes.xml with Word's separator notes (-1, 0) followed by the notes, numbered from 1. */
export function wordNotesXml(kind:'footnote'|'endnote',notes:{number:number;text:string}[]){
 const style=kind==='footnote'?'Footnote':'Endnote',separators=`<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;
 const body=notes.map(note=>`<w:${kind} w:id="${note.number}">${(note.text||'').split('\n').map((line,i)=>`<w:p><w:pPr><w:pStyle w:val="${style}Text"/></w:pPr>${i===0?`<w:r><w:rPr><w:rStyle w:val="${style}Reference"/></w:rPr><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`:''}${line?line.split('\t').map((part,j)=>(j?'<w:r><w:tab/></w:r>':'')+(part?`<w:r><w:t xml:space="preserve">${xmlText(part)}</w:t></w:r>`:'')).join(''):''}</w:p>`).join('')}</w:${kind}>`).join('');
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${separators}${body}</w:${kind}s>`;
}
/** Styles used by notes and the table of contents, added to the exported styles.xml. */
export const WORD_REFERENCE_STYLES=`<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>${[1,2,3].map(level=>`<w:style w:type="paragraph" w:styleId="TOC${level}"><w:name w:val="toc ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="${(level-1)*420}"/></w:pPr></w:style>`).join('')}<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>${['Footnote','Endnote'].map(name=>`<w:style w:type="paragraph" w:styleId="${name}Text"><w:name w:val="${name.toLowerCase()} text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="${name}Reference"><w:name w:val="${name.toLowerCase()} reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>`).join('')}`;
