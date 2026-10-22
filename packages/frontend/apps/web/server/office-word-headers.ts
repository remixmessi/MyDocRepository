import {posix} from 'node:path';
import {readOfficeXml} from './office-package';
import {decodeWordXml} from './office-word-link';
import {validWordHeaderFooter,WORD_PAGE_TOKEN,WORD_PAGES_TOKEN,wordHeaderParts,emptyWordLine,type WordHeaderFooter,type WordHeaderLine} from '../../../../common/nbstore/src/office-word-header';
import type {WordPageLayout} from '../../../../common/nbstore/src/office-word-page';

const attr=(xml:string,name:string)=>new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];
const fieldToken=(instr:string)=>{const name=instr.trim().split(/\s+/)[0]?.toUpperCase();return name==='PAGE'?WORD_PAGE_TOKEN:name==='NUMPAGES'||name==='SECTIONPAGES'?WORD_PAGES_TOKEN:undefined;};
const clean=(text:string)=>text.replace(new RegExp('[\\x00-\\x1f\\u2028\\u2029]','g'),' ').slice(0,255);
/**
 * One header/footer part as left/centre/right text. Tabs split the slots (Word's Header style has a centre and a
 * right tab); PAGE and NUMPAGES fields, simple or complex, become {页码}/{总页数}. Images, tables and other
 * content are not carried over.
 */
function readLine(xml:string):WordHeaderLine{
 const slots={left:[] as string[],center:[] as string[],right:[] as string[]};
 for(const paragraph of xml.matchAll(/<w:p(?:\s[^>]*?)?(?:\/>|>([\s\S]*?)<\/w:p>)/g)){
  const body=paragraph[1]??'',align=attr(/<w:jc\b[^>]*>/.exec(body)?.[0]??'','w:val'),content=body.replace(/<w:pPr\b[\s\S]*?<\/w:pPr>/g,'').replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g,'');
  const segments=[''];let state:'none'|'instr'|'result'='none',instr='',result='';
  const add=(text:string)=>{if(state==='result')result+=text;else if(state==='none')segments[segments.length-1]+=text;};
  for(const token of content.matchAll(/<w:fldSimple\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:fldSimple>)|<w:fldChar\b([^>]*?)\/?>|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:tab\b[^>]*\/>|<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)){
   const [whole,simpleAttrs,simpleBody,charAttrs,instrText,text]=token;
   if(simpleAttrs!==undefined){const value=fieldToken(decodeWordXml(attr(simpleAttrs,'w:instr')??''));add(value??[...(simpleBody??'').matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(m=>decodeWordXml(m[1])).join(''));}
   else if(charAttrs!==undefined){const type=attr(charAttrs,'w:fldCharType');if(type==='begin'){state='instr';instr='';result='';}else if(type==='separate')state='result';else if(type==='end'){state='none';add(fieldToken(instr)??result);}}
   else if(instrText!==undefined){if(state==='instr')instr+=decodeWordXml(instrText);}
   else if(whole.startsWith('<w:tab')){if(state==='none')segments.push('');}
   else if(text!==undefined)add(decodeWordXml(text));
  }
  const parts=segments.map(part=>part.trim());if(!parts.some(Boolean))continue;
  if(parts.length===1)(align==='center'?slots.center:align==='right'||align==='end'?slots.right:slots.left).push(parts[0]);
  else{slots.left.push(parts[0]);slots.center.push(parts[1]);if(parts.length>2)slots.right.push(parts.slice(2).filter(Boolean).join(' '));}
 }
 const join=(list:string[])=>clean(list.filter(Boolean).join(' '));
 return {left:join(slots.left),center:join(slots.center),right:join(slots.right)};
}
/** Default and first-page headers/footers of the single section; undefined when the document has none. */
export function readWordHeaderFooter(files:Record<string,Uint8Array>,documentXml:string):WordHeaderFooter|undefined{
 const sections=[...documentXml.replace(/<w:sectPrChange\b[^>]*>[\s\S]*?<\/w:sectPrChange>/g,'').matchAll(/<w:sectPr(?=\s|\/|>)(?:[^>]*\/\s*>|[^>]*>([\s\S]*?)<\/w:sectPr>)/g)],section=sections.at(-1)?.[1];if(!section)return;
 const relations=new Map<string,string>(),rels=files['word/_rels/document.xml.rels'];
 if(rels)for(const item of readOfficeXml(rels).matchAll(/<Relationship\b([^>]*)\/?>/g)){const id=attr(item[1],'Id'),target=attr(item[1],'Target');if(id&&target&&attr(item[1],'TargetMode')!=='External'&&!/[\\:#?]/.test(target)){const path=target.startsWith('/')?target.slice(1):posix.normalize(posix.join('word',target));if(path.startsWith('word/')&&!path.split('/').includes('..'))relations.set(id,path);}}
 const part=(kind:'header'|'footer',type:'default'|'first')=>{for(const ref of section.matchAll(new RegExp(`<w:${kind}Reference\\b([^>]*)\\/?>`,'g'))){if((attr(ref[1],'w:type')??'default')!==type)continue;const path=relations.get(attr(ref[1],'r:id')??''),bytes=path&&files[path];if(bytes)return readLine(readOfficeXml(bytes));}return {left:'',center:'',right:''};};
 const title=/<w:titlePg\b([^>]*)\/?>/.exec(section),firstPage=!!title&&!['0','false','off'].includes(attr(title[1],'w:val')??'1');
 const value:WordHeaderFooter={header:part('header','default'),footer:part('footer','default'),...(firstPage?{first:{header:part('header','first'),footer:part('footer','first')}}:{})};
 if(!firstPage&&emptyWordLine(value.header)&&emptyWordLine(value.footer))return;return validWordHeaderFooter(value)?value:undefined;
}

const xmlText=(value:string)=>value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
/** One header/footer part: a paragraph with centre and right tab stops at the middle and right edge of the text area. */
function partXml(kind:'hdr'|'ftr',line:WordHeaderLine,page:WordPageLayout){
 const width=page.width-page.left-page.right,runs=(text:string)=>wordHeaderParts(text).map(part=>'field' in part?`<w:fldSimple w:instr=" ${part.field} "><w:r><w:t>1</w:t></w:r></w:fldSimple>`:`<w:r><w:t xml:space="preserve">${xmlText(part.text)}</w:t></w:r>`).join('');
 const tab='<w:r><w:tab/></w:r>',body=runs(line.left)+(line.center||line.right?tab+runs(line.center):'')+(line.right?tab+runs(line.right):'');
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(width/2)}"/><w:tab w:val="right" w:pos="${width}"/></w:tabs></w:pPr>${body}</w:p></w:${kind}>`;
}
/** Parts, relationships, content types and sectPr references for export; empty when there is nothing to write. */
export function wordHeaderFooterParts(value:WordHeaderFooter|undefined,page:WordPageLayout){
 const files:Record<string,string>={},relations:string[]=[],types:string[]=[],references:string[]=[];if(!value)return {files,relations,types,references,titlePage:false};
 const add=(kind:'header'|'footer',type:'default'|'first',line:WordHeaderLine,n:number)=>{
  const name=`${kind}${n}.xml`,id=`rId${kind==='header'?'Hdr':'Ftr'}${n}`;files['word/'+name]=partXml(kind==='header'?'hdr':'ftr',line,page);
  relations.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${name}"/>`);
  types.push(`<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`);references.push(`<w:${kind}Reference w:type="${type}" r:id="${id}"/>`);
 };
 // A first-page variant is written even when empty, so a blank first page stays blank.
 if(!emptyWordLine(value.header))add('header','default',value.header,1);if(!emptyWordLine(value.footer))add('footer','default',value.footer,1);
 if(value.first){add('header','first',value.first.header,2);add('footer','first',value.first.footer,2);}
 // CT_SectPr: header references before footer references.
 references.sort((a,b)=>Number(b.startsWith('<w:header'))-Number(a.startsWith('<w:header')));
 return {files,relations,types,references,titlePage:!!value.first};
}
