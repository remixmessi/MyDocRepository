import {readOfficeXml} from './office-package';

// Only title level semantics are inherited, not arbitrary Word style properties.
function attribute(xml:string,name:string){return new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2].replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>');}
function property(xml:string,tag:string){return new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*?)(?:/>|>[\\s\\S]*?</w:${tag}>)`).exec(xml)?.[1];}
function paragraphProperties(xml:string){return /<w:pPr(?:\s[^>]*)?>([\s\S]*?)<\/w:pPr>/.exec(xml.replace(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/g,''))?.[1]??'';}
function outline(xml:string){const element=property(xml,'outlineLvl');if(element===undefined)return;const value=attribute(element,'w:val');if(value===undefined||! /^[0-9]$/.test(value))throw new Error('Word 大纲级别无效');return Number(value)<3?Number(value)+1:0;}
function builtIn(id:string){return /^Heading[123]$/i.test(id)?Number(id.slice(-1)):0;}
export function wordHeadingReader(bytes?:Uint8Array){
 const styles=new Map<string,{level:number|undefined;base:string|undefined}>(),resolved=new Map<string,{level:number;depth:number}>();
 if(bytes){
  const xml=readOfficeXml(bytes);
  for(const match of xml.matchAll(/<w:style(?=\s|>)([^>]*?)(?:\/>|>([\s\S]*?)<\/w:style>)/g)){
   if(attribute(match[1],'w:type')!=='paragraph')continue;
   const id=attribute(match[1],'w:styleId');if(!id)continue;
   if(styles.has(id))throw new Error('Word 段落样式编号重复');
   if(styles.size>=10000)throw new Error('Word 段落样式超过 1 万');
   const content=match[2]??'';
   styles.set(id,{level:outline(paragraphProperties(content)),base:attribute(property(content,'basedOn')??'','w:val')});
  }
 }
 const resolve=(id:string):number=>{
  const path:string[]=[],seen=new Set<string>();let current=id,level:number,depth=0;
  while(true){
   if(resolved.has(current)){const cached=resolved.get(current)!;if(path.length+cached.depth>64)throw new Error('Word 段落样式继承超过 64 层');level=cached.level;depth=cached.depth;break;}
   if(seen.has(current))throw new Error('Word 段落样式继承存在循环');
   if(path.length>=64)throw new Error('Word 段落样式继承超过 64 层');
   seen.add(current);path.push(current);const style=styles.get(current);
   if(!style){level=builtIn(current);break;}
   if(style.level!==undefined){level=style.level;break;}
   if(!style.base){level=builtIn(current);break;}
   current=style.base;
  }
  for(const key of path.reverse())resolved.set(key,{level,depth:++depth});return level;
 };
 for(const id of styles.keys())resolve(id);
 return (paragraph:string)=>{
  const properties=paragraphProperties(paragraph),direct=outline(properties);
  if(direct!==undefined)return direct;
  const id=attribute(property(properties,'pStyle')??'','w:val');return id?resolve(id):0;
 };
}
