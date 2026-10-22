import {validWordAlignment,validWordLeftIndent,validWordFirstLineIndent,validWordLineSpacing,type WordAlignment} from '../../../../common/nbstore/src/office-word-layout';
import {decodeWordXml} from './office-word-link';
import {readOfficeXml} from './office-package';
type Layout={pageBreakBefore?:boolean;align?:WordAlignment;leftIndent?:number;rightIndent?:number;firstLineIndent?:number;lineSpacing?:number;spaceBefore?:number;spaceAfter?:number};
function attribute(xml:string,name:string){const value=new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];return value===undefined?undefined:decodeWordXml(value);}
function property(xml:string,tag:string){return new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*?)(?:/>|>[\\s\\S]*?</w:${tag}>)`).exec(xml)?.[1];}
function properties(xml:string){return /<w:pPr(?:\s[^>]*)?>([\s\S]*?)<\/w:pPr>/.exec(xml.replace(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/g,''))?.[1]??'';}
function parse(xml:string):Layout{
 const layout:Layout={},jc=property(xml,'jc');
 const breakBefore=property(xml,'pageBreakBefore');if(breakBefore!==undefined){const value=attribute(breakBefore,'w:val');if(value!==undefined&&!['1','true','on','0','false','off'].includes(value))throw new Error('Word 段前分页值无效');layout.pageBreakBefore=value===undefined||['1','true','on'].includes(value);}
 if(jc!==undefined){const value=attribute(jc,'w:val'),align=value==='both'?'justify':value;if(!validWordAlignment(align))throw new Error('暂不支持该 Word 段落对齐方式');layout.align=align;}
 const ind=property(xml,'ind');
 if(ind!==undefined){
  if(['w:leftChars','w:rightChars','w:firstLineChars','w:hangingChars'].some(key=>attribute(ind,key)!==undefined))throw new Error('暂不支持字符单位的 Word 缩进');
  const twips=(value:string,label:string)=>{const number=Number(value);if(!/^\d+$/.test(value)||!validWordLeftIndent(number))throw new Error(`Word ${label}须在 0–144 磅范围内`);return number;};
  for(const [xmlKey,key,label]of [['w:left','leftIndent','左缩进'],['w:right','rightIndent','右缩进']] as const){const value=attribute(ind,xmlKey);if(value!==undefined)layout[key]=twips(value,label);}
  const hanging=attribute(ind,'w:hanging'),firstLine=attribute(ind,'w:firstLine');
  // OOXML gives hanging precedence when both are present.
  if(hanging!==undefined){const amount=twips(hanging,'悬挂缩进');layout.firstLineIndent=amount===0?0:-amount;}
  else if(firstLine!==undefined)layout.firstLineIndent=twips(firstLine,'首行缩进');
  if(layout.firstLineIndent!==undefined&&!validWordFirstLineIndent(layout.firstLineIndent))throw new Error('Word 特殊缩进无效');
 }
 // Spacing: the auto line rule is a multiple of 240; exact/at-least line heights and auto spacing stay at the default.
 const spacing=property(xml,'spacing');
 if(spacing!==undefined){
  const number=(key:string)=>{const value=attribute(spacing,key);return value!==undefined&&/^\d+$/.test(value)?Number(value):undefined;};
  const line=number('w:line'),rule=attribute(spacing,'w:lineRule')??'auto';if(line!==undefined&&rule==='auto'){const multiple=Math.round(line/240*100)/100;if(validWordLineSpacing(multiple))layout.lineSpacing=multiple;}
  for(const [key,field,auto]of [['w:before','spaceBefore','w:beforeAutospacing'],['w:after','spaceAfter','w:afterAutospacing']] as const){const value=number(key);if(value!==undefined&&!['1','true','on'].includes(attribute(spacing,auto)??''))layout[field]=Math.min(value,1440);}
 }

 return layout;
}
export function wordLayoutReader(bytes?:Uint8Array){
 const styles=new Map<string,{layout:Layout;base?:string}>(),resolved=new Map<string,{layout:Layout;depth:number}>();let defaultStyle:string|undefined;
 if(bytes)for(const match of readOfficeXml(bytes).matchAll(/<w:style(?=\s|>)([^>]*?)(?:\/>|>([\s\S]*?)<\/w:style>)/g)){
  if(attribute(match[1],'w:type')!=='paragraph')continue;const id=attribute(match[1],'w:styleId');if(!id)continue;
  if(styles.has(id))throw new Error('Word 段落样式编号重复');if(styles.size>=10000)throw new Error('Word 段落样式过多');
  const content=match[2]??'';styles.set(id,{layout:parse(properties(content)),base:attribute(property(content,'basedOn')??'','w:val')});
  if(['1','true','on'].includes(attribute(match[1],'w:default')??''))defaultStyle=id;
 }
 const resolve=(id:string):Layout=>{
  const path:string[]=[],seen=new Set<string>();let current:string|undefined=id,base:Layout={},depth=0;
  while(current){
   const cached=resolved.get(current);if(cached){if(path.length+cached.depth>64)throw new Error('Word 段落样式继承超过 64 层');base=cached.layout;depth=cached.depth;break;}
   if(seen.has(current)||path.length>=64)throw new Error('Word 段落样式继承无效');seen.add(current);path.push(current);current=styles.get(current)?.base;
  }
  for(const key of path.reverse()){base={...base,...styles.get(key)?.layout};resolved.set(key,{layout:base,depth:++depth});}return base;
 };
 for(const id of styles.keys())resolve(id);
 return (paragraph:string):Layout=>{const p=properties(paragraph),style=attribute(property(p,'pStyle')??'','w:val')??defaultStyle;return {...(style?resolve(style):{}),...parse(p)};};
}
