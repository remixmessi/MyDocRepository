import {readingJump} from './document-reading';
import type { Node } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
export const MAX_FIND_QUERY=256;
export const MAX_FIND_MATCHES=5000;
export interface FindMatch { from:number; to:number; codeFrom?:number; codeTo?:number }
export interface DocumentFindState { open:boolean; query:string; caseSensitive:boolean; matches:FindMatch[]; index:number; truncated:boolean }
interface FindPluginState extends DocumentFindState { decorations:DecorationSet }
interface FindAction { open?:boolean; query?:string; caseSensitive?:boolean; direction?:1|-1 }
export const emptyDocumentFind:DocumentFindState={open:false,query:'',caseSensitive:false,matches:[],index:-1,truncated:false};
export function findDocumentMatches(doc:Node,query:string,caseSensitive=false,limit=MAX_FIND_MATCHES){
  const matches:FindMatch[]=[];let truncated=false;query=query.slice(0,MAX_FIND_QUERY);
  if(!query)return {matches,truncated};
  const expression=new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),caseSensitive?'gu':'giu');
  doc.descendants((node,pos)=>{
    if(truncated)return false;if(!node.isTextblock)return;
    // Inline atoms are one document position; a sentinel preserves offsets and prevents cross-atom matches.
    const text=node.textBetween(0,node.content.size,'',leaf=>leaf.type.name==='hard_break'?'\n':'\ufffc');
    expression.lastIndex=0;let match:RegExpExecArray|null;
    while((match=expression.exec(text))){if(matches.length>=limit){truncated=true;break;}matches.push({from:pos+1+match.index,to:pos+1+match.index+match[0].length,...(node.type.name==='code_block'?{codeFrom:pos,codeTo:pos+node.nodeSize}:{})});}
    return false;
  });return {matches,truncated};
}
export const documentFindKey=new PluginKey<FindPluginState>('document-find');
function decorate(doc:Node,state:DocumentFindState){
  const decorations:Decoration[]=[];const codeBlocks=new Map<number,number>();
  state.matches.forEach((match,index)=>{decorations.push(Decoration.inline(match.from,match.to,{class:'document-find-match'+(index===state.index?' document-find-current':''),'data-find-index':String(index)}));if(match.codeFrom!==undefined)codeBlocks.set(match.codeFrom,match.codeTo!);});
  for(const [from,to]of codeBlocks)decorations.push(Decoration.node(from,to,{class:'document-find-reveal'}));
  return DecorationSet.create(doc,decorations);
}
export function updateDocumentFind(view:EditorView,action:FindAction){view.dispatch(view.state.tr.setMeta(documentFindKey,action).setMeta('addToHistory',false));}
export function createDocumentFind(notify:(state:DocumentFindState)=>void){return new Plugin<FindPluginState>({
  key:documentFindKey,
  state:{init:()=>({...emptyDocumentFind,decorations:DecorationSet.empty}),apply(tr,previous){
    const action=tr.getMeta(documentFindKey) as FindAction|undefined;if(!action&&(!tr.docChanged||!previous.open))return previous;
    const query=(action?.query??previous.query).slice(0,MAX_FIND_QUERY),open=action?.open??previous.open,caseSensitive=action?.caseSensitive??previous.caseSensitive;
    const result=open?(!tr.docChanged&&previous.open&&query===previous.query&&caseSensitive===previous.caseSensitive?{matches:previous.matches,truncated:previous.truncated}:findDocumentMatches(tr.doc,query,caseSensitive)):{matches:[],truncated:false};
    let index=result.matches.length?0:-1;
    if(index>=0&&query===previous.query&&caseSensitive===previous.caseSensitive){
      const current=previous.matches[previous.index];if(current){const mapped=tr.mapping.map(current.from);index=result.matches.findIndex(m=>m.from>=mapped);if(index<0)index=result.matches.length-1;}
    }
    if(action?.direction&&result.matches.length)index=(index+action.direction+result.matches.length)%result.matches.length;
    const next={open,query,caseSensitive,...result,index};return {...next,decorations:decorate(tr.doc,next)};
  }},
  props:{decorations:state=>documentFindKey.getState(state)?.decorations},
  view(view){let frame=0;notify(documentFindKey.getState(view.state)!);return {update(view,previousState){const next=documentFindKey.getState(view.state)!,old=documentFindKey.getState(previousState)!;if(next===old)return;notify(next);cancelAnimationFrame(frame);if(next.open&&next.index>=0&&(next.query!==old.query||next.index!==old.index||next.caseSensitive!==old.caseSensitive||!old.open)&&!view.hasFocus())frame=requestAnimationFrame(()=>{if(view.isDestroyed)return;(readingJump(view),view.dom.querySelector('.document-find-current'))?.scrollIntoView({block:'center',inline:'nearest'});});},destroy(){cancelAnimationFrame(frame);}};},
});}
