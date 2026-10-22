import {readOfficeXml} from './office-package';
import type {WordList} from '../../../../common/nbstore/src/office-word-list';
function attribute(xml:string,name:string){return new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2].replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi,(_,entity:string)=>entity[0]==='#'?String.fromCodePoint(entity[1].toLowerCase()==='x'?parseInt(entity.slice(2),16):Number(entity.slice(1))):({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"} as Record<string,string>)[entity]);}
function value(xml:string,tag:string){const attributes=new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*)>`).exec(xml)?.[1];return attributes===undefined?undefined:attribute(attributes,'w:val');}
function integer(input:string|undefined,name:string,min=0,max=1000000){if(input===undefined||!/^\d+$/.test(input)||Number(input)<min||Number(input)>max)throw new Error('Word 列表'+name+'无效');return Number(input);}
export function wordListReader(bytes?:Uint8Array,styleBytes?:Uint8Array){
 const abstracts=new Map<number,{kind:WordList['kind'];start:number}>(),numbers=new Map<number,WordList>();
 if(bytes){
  const xml=readOfficeXml(bytes),abstractIds=new Set<number>(),numberIds=new Set<number>();
  for(const match of xml.matchAll(/<w:abstractNum(?=\s|>)([^>]*)>([\s\S]*?)<\/w:abstractNum>/g)){
   const id=integer(attribute(match[1],'w:abstractNumId'),'定义编号');if(abstractIds.has(id))throw new Error('Word 列表定义编号重复');if(abstractIds.size>=10000)throw new Error('Word 列表定义过多');abstractIds.add(id);
   const level=[...match[2].matchAll(/<w:lvl(?=\s|>)([^>]*)>([\s\S]*?)<\/w:lvl>/g)].find(item=>attribute(item[1],'w:ilvl')==='0');
   if(!level)continue;const format=value(level[2],'numFmt');if(format!=='bullet'&&format!=='decimal')continue;
   const pattern=value(level[2],'lvlText');
   if(format==='decimal'&&pattern!=='%1.'||format==='bullet'&&!['•','●','\uf0b7'].includes(pattern??''))continue;
   abstracts.set(id,{kind:format,start:integer(value(level[2],'start')??'1','起始值',1)});
  }
  for(const match of xml.matchAll(/<w:num(?=\s|>)([^>]*)>([\s\S]*?)<\/w:num>/g)){
   const id=integer(attribute(match[1],'w:numId'),'实例编号');if(numberIds.has(id))throw new Error('Word 列表实例编号重复');if(numberIds.size>=10000)throw new Error('Word 列表实例过多');numberIds.add(id);
   const definition=abstracts.get(integer(value(match[2],'abstractNumId'),'关联编号'));if(!definition)continue;
   let start=definition.start;
   for(const override of match[2].matchAll(/<w:lvlOverride(?=\s|>)([^>]*)>([\s\S]*?)<\/w:lvlOverride>/g)){
    if(attribute(override[1],'w:ilvl')!=='0')continue;
    if(/<w:lvl(?=\s|>)/.test(override[2]))throw new Error('暂不支持 Word 列表层级格式覆盖');
    const explicit=value(override[2],'startOverride');if(explicit!==undefined)start=integer(explicit,'起始值',1);
   }
   numbers.set(id,{kind:definition.kind,id:'import-'+id,start});
  }
 }
 const properties=(xml:string)=>/<w:pPr(?:\s[^>]*)?>([\s\S]*?)<\/w:pPr>/.exec(xml.replace(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/g,''))?.[1]??'';
 const parse=(xml:string):WordList|null|undefined=>{
  const numbering=/<w:numPr(?:\s[^>]*)?>([\s\S]*?)<\/w:numPr>/.exec(xml)?.[1];if(!numbering)return;
  const id=integer(value(numbering,'numId'),'段落编号');if(id===0)return null;
  if(integer(value(numbering,'ilvl')??'0','层级',0,8)!==0)throw new Error('暂不支持多级 Word 列表，请使用单层列表');
  const list=numbers.get(id);if(!list)throw new Error('Word 列表定义缺失或编号格式尚不支持');return {...list};
 };
 const styles=new Map<string,{properties:string;base:string|undefined}>();
 if(styleBytes)for(const match of readOfficeXml(styleBytes).matchAll(/<w:style(?=\s|>)([^>]*?)(?:\/>|>([\s\S]*?)<\/w:style>)/g)){
  if(attribute(match[1],'w:type')!=='paragraph')continue;const id=attribute(match[1],'w:styleId');if(!id)continue;
  if(styles.size>=10000)throw new Error('Word 段落样式过多');styles.set(id,{properties:properties(match[2]??''),base:value(match[2]??'','basedOn')});
 }
 return (paragraph:string):WordList|undefined=>{
  const paragraphProperties=properties(paragraph),direct=parse(paragraphProperties);if(direct!==undefined)return direct??undefined;
  let id=value(paragraphProperties,'pStyle');const seen=new Set<string>();
  while(id){
   if(seen.has(id)||seen.size>=64)throw new Error('Word 列表样式继承无效');seen.add(id);const style=styles.get(id);if(!style)break;
   const inherited=parse(style.properties);if(inherited!==undefined)return inherited??undefined;id=style.base;
  }
 };
}
