import {validWordFontSize,validWordColor} from '../../../../common/nbstore/src/office-word-font';
import {decodeWordXml} from './office-word-link';
import {readOfficeXml} from './office-package';
export type WordRunMarks={bold?:boolean;italic?:boolean;underline?:boolean;strike?:boolean;fontSize?:number;color?:string};
const TOGGLES=[['b','bold'],['i','italic'],['strike','strike']] as const;
function attribute(xml:string,name:string){const value=new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];return value===undefined?undefined:decodeWordXml(value);}
function property(xml:string,tag:string){return new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*?)(?:/>|>[\\s\\S]*?</w:${tag}>)`).exec(xml)?.[1];}
export function wordRunProperties(xml:string){return /<w:rPr(?:\s[^>]*)?>([\s\S]*?)<\/w:rPr>/.exec(xml.replace(/<w:rPrChange\b[^>]*>[\s\S]*?<\/w:rPrChange>/g,''))?.[1]??'';}
/** Style definitions are read leniently: values the editor cannot represent are ignored rather than rejecting the file. */
function parseStyleMarks(rPr:string):WordRunMarks{
 const marks:WordRunMarks={};
 for(const [tag,key]of TOGGLES){const element=property(rPr,tag);if(element===undefined)continue;const value=attribute(element,'w:val');if(value===undefined||['1','true','on'].includes(value))marks[key]=true;else if(['0','false','off'].includes(value))marks[key]=false;}
 const u=property(rPr,'u');if(u!==undefined){const value=attribute(u,'w:val');if(value==='single')marks.underline=true;else if(value==='none')marks.underline=false;}
 const sz=property(rPr,'sz');if(sz!==undefined){const value=attribute(sz,'w:val'),points=Number(value)/2;if(value&&/^\d+$/.test(value)&&validWordFontSize(points))marks.fontSize=points;}
 const color=property(rPr,'color');if(color!==undefined&&!/\bw:theme(?:Color|Tint|Shade)\s*=/.test(color)){const value=attribute(color,'w:val'),normalized=value==='auto'?'auto':'#'+(value??'').toUpperCase();if(validWordColor(normalized))marks.color=normalized;}
 return marks;
}
// Bold and italic are toggle properties across style levels (ECMA-376 17.7.3); other marks override.
function layer(base:WordRunMarks,next:WordRunMarks):WordRunMarks{
 const result={...base,...next};for(const [,key]of TOGGLES)if(next[key]!==undefined)result[key]=!!base[key]!==!!next[key];return result;
}
export function wordRunStyleReader(bytes?:Uint8Array){
 const styles=new Map<string,{type:string;marks:WordRunMarks;base?:string}>();let defaults:WordRunMarks={},defaultParagraph:string|undefined;
 if(bytes){
  const xml=readOfficeXml(bytes),docDefaults=/<w:rPrDefault(?:\s[^>]*)?>([\s\S]*?)<\/w:rPrDefault>/.exec(xml)?.[1];if(docDefaults)defaults=parseStyleMarks(wordRunProperties(docDefaults));
  for(const match of xml.matchAll(/<w:style(?=\s|>)([^>]*?)(?:\/>|>([\s\S]*?)<\/w:style>)/g)){
   const type=attribute(match[1],'w:type')??'',id=attribute(match[1],'w:styleId');if(!id||!['paragraph','character'].includes(type))continue;
   const key=type+':'+id;if(styles.has(key))throw new Error('Word 样式编号重复');if(styles.size>=20000)throw new Error('Word 样式过多');
   const content=match[2]??'';styles.set(key,{type,marks:parseStyleMarks(wordRunProperties(content)),base:attribute(property(content,'basedOn')??'','w:val')});
   if(type==='paragraph'&&['1','true','on'].includes(attribute(match[1],'w:default')??''))defaultParagraph=id;
  }
 }
 const cache=new Map<string,WordRunMarks>();
 const chain=(type:'paragraph'|'character',id:string|undefined):WordRunMarks=>{
  if(!id)return {};const key=type+':'+id,cached=cache.get(key);if(cached)return cached;
  const path:string[]=[],seen=new Set<string>();let current:string|undefined=id;
  while(current&&styles.has(type+':'+current)){if(seen.has(current)||path.length>=64)throw new Error('Word 样式继承无效');seen.add(current);path.push(current);current=styles.get(type+':'+current)!.base;}
  const marks=path.reverse().reduce<WordRunMarks>((value,name)=>layer(value,styles.get(type+':'+name)!.marks),{});cache.set(key,marks);return marks;
 };
 const baseline=layer(defaults,chain('paragraph',defaultParagraph));
 /** Effective inherited marks that differ from the document's normal text; direct run formatting is applied by the caller. */
 return (paragraphStyle:string|undefined,runProperties:string,heading:boolean):WordRunMarks=>{
  const paragraph=layer(defaults,chain('paragraph',paragraphStyle??defaultParagraph)),inherited=layer(paragraph,chain('character',attribute(property(runProperties,'rStyle')??'','w:val')));
  const result:WordRunMarks={};
  for(const key of ['bold','italic','underline','strike'] as const)if(inherited[key]&&!(heading&&key==='bold'&&paragraph.bold))result[key]=true;
  // Headings already render their level's size; normal-text size and colour stay implicit as before.
  if(inherited.fontSize!==undefined&&inherited.fontSize!==baseline.fontSize&&!(heading&&inherited.fontSize===paragraph.fontSize))result.fontSize=inherited.fontSize;
  if(inherited.color!==undefined&&inherited.color!==baseline.color)result.color=inherited.color;
  return result;
 };
}
