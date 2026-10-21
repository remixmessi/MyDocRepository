import {readOfficeXml} from './office-package';
export type XlsxXml={name:string;attributes:Record<string,string>;children:XlsxXml[];text:string};
export function xmlValue(value:string){return value.replace(/&([^;\s<&]+);|&/g,(_,entity:string|undefined)=>{if(entity===undefined)throw new Error('XLSX XML 实体无效');const values:Record<string,string>={lt:'<',gt:'>',amp:'&',quot:'"',apos:"'"};if(Object.hasOwn(values,entity))return values[entity];if(!/^#(?:x[0-9a-f]+|[0-9]+)$/i.test(entity))throw new Error('XLSX XML 实体无效');const n=entity[1].toLowerCase()==='x'?parseInt(entity.slice(2),16):Number(entity.slice(1));if(!(n===9||n===10||n===13||n>=32&&n<=0xd7ff||n>=0xe000&&n<=0xfffd||n>=0x10000&&n<=0x10ffff))throw new Error('XLSX XML 字符无效');return String.fromCodePoint(n);});}
export function xlsxXml(bytes:Uint8Array){
 const xml=readOfficeXml(bytes).replace(/\r\n|\r/g,'\n'),stack:XlsxXml[]=[],roots:XlsxXml[]=[];
 for(const match of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:[^>"']|"[^"]*"|'[^']*')*>|[^<]+/g)){
  const token=match[0];if(token.startsWith('<!--')||token.startsWith('<?'))continue;
  if(token.startsWith('<![CDATA[')){if(!stack.length)throw new Error('XLSX XML 文本位置无效');stack.at(-1)!.text+=token.slice(9,-3);continue;}
  if(token.startsWith('</')){stack.pop();continue;}
  if(token.startsWith('<')){
   const found=/^<([\w.:-]+)/.exec(token);if(!found)throw new Error('XLSX XML 标签无效');const attributes:Record<string,string>=Object.create(null);let rest=token.slice(found[0].length).replace(/\/?\s*>$/,'');
   while(rest.trim()){const attr=/^\s+([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(rest);if(!attr||Object.hasOwn(attributes,attr[1]))throw new Error('XLSX XML 属性无效');if((attr[2]??attr[3]).includes('<'))throw new Error('XLSX XML 属性无效');attributes[attr[1]]=xmlValue(attr[2]??attr[3]);rest=rest.slice(attr[0].length);}
   const node:XlsxXml={name:found[1].split(':').at(-1)!,attributes,children:[],text:''};if(stack.length)stack.at(-1)!.children.push(node);else roots.push(node);if(!/\/\s*>$/.test(token))stack.push(node);
  }else if(stack.length)stack.at(-1)!.text+=xmlValue(token);else if(token.trim())throw new Error('XLSX XML 根外存在文本');
 }
 if(roots.length!==1||stack.length)throw new Error('XLSX XML 结构无效');return roots[0];
}
export function attribute(node:XlsxXml,name:string){const matches=Object.entries(node.attributes).filter(([key])=>key.split(':').at(-1)===name);if(matches.length>1)throw new Error('XLSX XML 属性存在歧义');return matches[0]?.[1];}
export function children(node:XlsxXml,name:string){return node.children.filter(child=>child.name===name);}
export function child(node:XlsxXml,name:string){const found=children(node,name);if(found.length>1)throw new Error('XLSX XML 部件重复：'+name);return found[0];}
export function stringValue(node:XlsxXml){
 const decodeNode=(text:XlsxXml)=>{if(text.children.length)throw new Error('XLSX 文本结构无效');return decode(text.text);};
 const decode=(value:string)=>value.replace(/_x([0-9a-f]{4})_/gi,(_,hex:string)=>String.fromCharCode(parseInt(hex,16)));
 return node.children.flatMap(part=>part.name==='t'?[decodeNode(part)]:part.name==='r'?children(part,'t').map(decodeNode):[]).join('');
}
