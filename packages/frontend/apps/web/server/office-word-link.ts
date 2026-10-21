import {validWordLink} from '../../../../common/nbstore/src/office-word-link';
import {readOfficeXml} from './office-package';
export function decodeWordXml(value:string){return value.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi,(_,entity:string)=>{
 if(entity.startsWith('#')){const code=entity[1].toLowerCase()==='x'?parseInt(entity.slice(2),16):Number(entity.slice(1));if(code>0x10ffff||code>=0xd800&&code<=0xdfff)throw new Error('XML 字符无效');return String.fromCodePoint(code);}
 return ({lt:'<',gt:'>',amp:'&',quot:'"',apos:"'"} as Record<string,string>)[entity];
});}
function attribute(xml:string,name:string){const value=new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];return value===undefined?undefined:decodeWordXml(value);}
export function wordLinkReader(bytes?:Uint8Array){
 const relations=new Map<string,{target:string|undefined;type:string|undefined;mode:string|undefined}>();
 if(bytes)for(const match of readOfficeXml(bytes).matchAll(/<(?:[\w.-]+:)?Relationship(?=\s|>)([^>]*?)(?:\/>|>\s*<\/(?:[\w.-]+:)?Relationship>)/g)){
  const id=attribute(match[1],'Id');if(!id)throw new Error('Word 关系编号缺失');
  if(relations.has(id))throw new Error('Word 关系编号重复');if(relations.size>=10000)throw new Error('Word 关系数量超过 1 万');
  relations.set(id,{target:attribute(match[1],'Target'),type:attribute(match[1],'Type'),mode:attribute(match[1],'TargetMode')});
 }
 return (attributes:string)=>{
  if(attribute(attributes,'w:anchor')!==undefined)throw new Error('暂不支持 Word 书签链接');
  const id=attribute(attributes,'r:id'),relation=id?relations.get(id):undefined;
  if(!relation||relation.type!=='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'||relation.mode!=='External')throw new Error('Word 超链接关系缺失或类型不支持');
  if(!validWordLink(relation.target))throw new Error('Word 仅支持无账号密码的 HTTP/HTTPS 网页链接');
  return relation.target;
 };
}
/** Keep run order and distinguish linked runs from adjacent ordinary runs. */
export function* wordRuns(paragraph:string,linkFor:(attributes:string)=>string):Generator<{xml:string;href?:string}>{
 if(/<w:(?:fldSimple|instrText)(?=\s|>)/.test(paragraph))throw new Error('暂不支持 Word 域代码，请转换为普通文字或网页链接');
 for(const token of paragraph.matchAll(/<w:hyperlink(?=\s|>)([^>]*?)(?:\/>|>([\s\S]*?)<\/w:hyperlink>)|<w:r(?:\s[^>]*?)?>([\s\S]*?)<\/w:r>/g)){
  if(token[3]!==undefined){yield {xml:token[3]};continue;}
  const href=linkFor(token[1]);if(/<w:hyperlink(?=\s|>)/.test(token[2]??''))throw new Error('Word 超链接不可嵌套');
  for(const run of (token[2]??'').matchAll(/<w:r(?:\s[^>]*?)?>([\s\S]*?)<\/w:r>/g))yield {xml:run[1],href};
 }
}
