import type MarkdownIt from 'markdown-it';
import footnote from 'markdown-it-footnote';
import type {NodeSpec,Node as PMNode} from 'prosemirror-model';
import type {MarkdownSerializerState} from 'prosemirror-markdown';
export const validFootnoteLabel=(value:unknown):value is string=>typeof value==='string'&&value.length>0&&value.length<=160&&!/[\s\[\]\u0000-\u001f]/.test(value);
export const footnoteNodes:Record<string,NodeSpec>={
 footnote_reference:{inline:true,group:'inline',atom:true,attrs:{label:{}},parseDOM:[{tag:'sup[data-footnote-reference]',getAttrs:el=>{const label=(el as HTMLElement).dataset.footnoteReference;return validFootnoteLabel(label)?{label}:false;}}],toDOM:node=>['sup',{'data-footnote-reference':node.attrs.label},['a',{href:'#footnote-'+encodeURIComponent(node.attrs.label)},'脚注']]},
 footnote_definition:{group:'block',content:'block+',defining:true,attrs:{label:{}},parseDOM:[{tag:'section[data-footnote-definition]',getAttrs:el=>{const label=(el as HTMLElement).dataset.footnoteDefinition;return validFootnoteLabel(label)?{label}:false;}}],toDOM:node=>['section',{'data-footnote-definition':node.attrs.label,class:'footnote-definition'},0]},
};
export function footnoteIndex(doc:PMNode){const definitions=new Map<string,{node:PMNode;position:number}>(),references=new Map<string,number[]>(),numbers=new Map<string,number>();doc.descendants((node,position)=>{if(node.type.name==='footnote_definition'&&!definitions.has(node.attrs.label))definitions.set(node.attrs.label,{node,position});if(node.type.name==='footnote_reference'){const label=node.attrs.label;if(!numbers.has(label))numbers.set(label,numbers.size+1);const positions=references.get(label)??[];positions.push(position);references.set(label,positions);}});return {definitions,references,numbers};}
export function footnoteMarkdown(md:InstanceType<typeof MarkdownIt>){md.use(footnote);
 // Preserve orphan definitions instead of dropping them in the dependency's HTML-oriented tail pass.
 md.core.ruler.at('footnote_tail',state=>{
  const defs:{label:string;tokens:typeof state.tokens}[]=[],body:typeof state.tokens=[];let current:typeof defs[number]|undefined;
  for(const token of state.tokens){if(token.type==='footnote_reference_open'){if(current)throw new Error('不支持嵌套脚注定义');const label=token.meta?.label;if(!validFootnoteLabel(label))throw new Error('脚注标识无效或超过 160 字符');current={label,tokens:[]};defs.push(current);}else if(token.type==='footnote_reference_close')current=undefined;else (current?current.tokens:body).push(token);}
  const labels=new Set(defs.map(def=>def.label));const env=state.env as {footnotes?:{list?:{label?:string;tokens?:typeof state.tokens}[]}};const list=env.footnotes?.list??[];
  for(let i=0;i<list.length;i++){const item=list[i];if(item.label)continue;let label=`fn-${i+1}`;while(labels.has(label))label+='-';labels.add(label);item.label=label;const open=new state.Token('paragraph_open','p',1),inline=new state.Token('inline','',0),close=new state.Token('paragraph_close','p',-1);inline.children=item.tokens??[];defs.push({label,tokens:[open,inline,close]});}
  let refs=0;const visit=(tokens:typeof state.tokens)=>{for(const token of tokens){if(token.type==='footnote_ref'){const meta=token.meta??{};meta.label=meta.label??(typeof meta.id==='number'?list[meta.id]?.label:undefined);token.meta=meta;if(!validFootnoteLabel(meta.label))throw new Error('脚注标识无效或超过 160 字符');if(++refs>5000)throw new Error('每篇最多 5000 处脚注引用');}if(token.children)visit(token.children);}};visit(body);for(const def of defs)visit(def.tokens);
  if(defs.length>500)throw new Error('每篇最多 500 个脚注');
  const used=new Set<string>();for(const def of defs){if(!validFootnoteLabel(def.label))throw new Error('脚注标识无效或超过 160 字符');let label=def.label;let suffix=2;while(used.has(label))label=def.label.slice(0,150)+'-'+suffix++;used.add(label);const open=new state.Token('footnote_definition_open','section',1);open.attrSet('label',label);body.push(open,...def.tokens,new state.Token('footnote_definition_close','section',-1));}state.tokens=body;
 });
}
export const footnoteSerializers={
 footnote_reference(state:MarkdownSerializerState,node:PMNode){state.write(`[^${node.attrs.label}]`);},
 footnote_definition(state:MarkdownSerializerState,node:PMNode){state.write(`[^${node.attrs.label}]: `);state.wrapBlock('    ',null,node,()=>state.renderContent(node));},
};
