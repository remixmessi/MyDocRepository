import type{Node as PMNode}from'prosemirror-model';
import{markdownParser}from'../document-format';import{documentText}from'../document-links';import type{Shape,Point}from'./model';
export interface BoardNote extends Omit<Shape,'kind'>{kind:'note';markdown:string}
const allowed=new Set(['doc','callout','math_inline','math_block','mermaid','doc_reference','paragraph','text','heading','blockquote','code_block','horizontal_rule','bullet_list','ordered_list','list_item','hard_break','table','table_row','table_cell','table_header']);
const cache=new Map<string,PMNode>();
export function parseNote(markdown:string){if(typeof markdown!=='string'||markdown.length>20000)throw new Error('每篇白板笔记最多支持 20,000 个 Markdown 字符');const cached=cache.get(markdown);if(cached)return cached;const doc=markdownParser.parse(markdown);let count=0;doc.descendants((node,pos)=>{count++;if(!allowed.has(node.type.name))throw new Error('白板笔记暂不支持图片、附件或文档嵌入');if(count>1000||doc.resolve(pos).depth>40)throw new Error('笔记结构过大或层级过深');});if(cache.size>=128)cache.delete(cache.keys().next().value!);cache.set(markdown,doc);return doc;}
export const noteText=(markdown:string,resolve?:(id:string,workspaceId:string)=>string|undefined)=>documentText(parseNote(markdown),resolve);
export function validNote(value:unknown):value is BoardNote{const n=value as BoardNote;if(!n||n.kind!=='note'||n.width<240||n.height<160||typeof n.markdown!=='string')return false;try{return n.text===noteText(n.markdown).slice(0,2000);}catch{return false;}}
export function createNote(point:Point):BoardNote{return{id:crypto.randomUUID(),kind:'note',...point,width:360,height:320,text:'',markdown:'',fill:'#ffffff'};}
