import {useEffect,useState,type RefObject} from 'react';
import {MessageSquare} from 'lucide-react';
import {Y,documentAwareness} from '@mydoc/crdt';
import {officeAnchorTarget} from '../../../common/nbstore/src/office-comment-anchor';
import {validOfficeComment,type OfficeCommentThread} from '../../../common/nbstore/src/office-comments';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import type {OfficePeer} from '../../../common/nbstore/src/office-presence';
import {wordTextIdentity} from './office-word-flow';
import {wordDomPoint} from './word-flow-dom';
import {requestOfficeThread} from './office-collaboration-ui';

type Box={left:number;top:number;width:number;height:number};
type Layer={carets:{client:number;name:string;color:string;box:Box}[];selections:{client:number;color:string;boxes:Box[]}[];comments:{boxes:Box[]}[];badges:{top:number;threads:string[];label:string}[];margin:number};
type Resolved={text:Y.Text;start:number;end:number};
const MAX_BOXES=200;

/**
 * Remote carets, remote selections and unresolved comment ranges drawn over the Word paper. Positions come from
 * the browser's own line boxes, converted back into the paper's unzoomed coordinates, so they follow wrapping,
 * zoom and pagination. The layer is decorative (aria-hidden) except the comment badges, which open the thread.
 */
export function WordCollaborationLayer({paper,doc,workspace,id,peers,resolve}:{paper:RefObject<HTMLElement>;doc:Y.Doc;workspace:string;id:string;peers:OfficePeer[];resolve:(peer:OfficePeer)=>Resolved|undefined}){
 const [layer,setLayer]=useState<Layer>({carets:[],selections:[],comments:[],badges:[],margin:0}),peerKey=JSON.stringify(peers.map(peer=>[peer.client,peer.color,peer.name,peer.focus]));
 useEffect(()=>{
  const element=paper.current;if(!element)return;let frame=0;
  const measure=()=>{cancelAnimationFrame(frame);frame=requestAnimationFrame(()=>{
   const origin=element.getBoundingClientRect(),zoom=origin.width/(element.offsetWidth||1)||1;
   const box=(rect:DOMRect):Box=>({left:(rect.left-origin.left)/zoom,top:(rect.top-origin.top)/zoom,width:rect.width/zoom,height:rect.height/zoom});
   const input=(text:Y.Text)=>{const identity=wordTextIdentity(text);return identity?element.querySelector<HTMLElement>(`.word-rich-input[data-word-text="${identity}"]`):null;};
   const range=(node:HTMLElement,start:number,end:number)=>{const value=document.createRange();value.setStart(...wordDomPoint(node,start));value.setEnd(...wordDomPoint(node,end));return value;};
   const boxes=(value:Range)=>[...value.getClientRects()].filter(rect=>rect.width>0.5||rect.height>0.5).slice(0,MAX_BOXES).map(box);
   // Badges sit in the right page margin, like Word's comment balloons.
   const next:Layer={carets:[],selections:[],comments:[],badges:[],margin:Math.max(0,element.offsetWidth-34)};
   for(const peer of peers){const target=resolve(peer),node=target&&input(target.text);if(!target||!node)continue;
    if(target.end>target.start)next.selections.push({client:peer.client,color:peer.color,boxes:boxes(range(node,target.start,target.end))});
    // A collapsed range has no width; measure the character before (or after) the caret and take its edge.
    const length=target.text.length,at=target.end,before=at>0,probe=before?range(node,at-1,at):range(node,at,Math.min(length,at+1)),rect=[...probe.getClientRects()].pop();
    const caret=rect?{left:(before?rect.right:rect.left),top:rect.top,height:rect.height}:(()=>{const fallback=node.getBoundingClientRect();return {left:fallback.left,top:fallback.top,height:Math.min(fallback.height,24)||18};})();
    next.carets.push({client:peer.client,name:peer.name,color:peer.color,box:{left:(caret.left-origin.left)/zoom,top:(caret.top-origin.top)/zoom,width:2,height:caret.height/zoom}});
   }
   const texts=wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray()),threads=doc.share.has('officeComments')?[...doc.getMap<OfficeCommentThread>('officeComments').values()].filter(thread=>validOfficeComment(thread)&&!thread.resolved&&thread.anchor?.kind==='text'):[];
   const badges=new Map<number,{threads:string[];label:string}>();
   for(const thread of threads){const target=officeAnchorTarget(doc,thread.anchor!);if(target?.kind!=='text')continue;const first=texts.indexOf(target.first),last=texts.indexOf(target.last);if(first<0||last<first)continue;
    const found:Box[]=[];for(let i=first;i<=last&&found.length<MAX_BOXES;i++){const node=input(texts[i]);if(!node)continue;found.push(...boxes(range(node,i===first?target.start:0,i===last?target.end:texts[i].length)));}
    if(!found.length)continue;next.comments.push({boxes:found});
    // Badges within one line of each other share a button, like Word's margin balloons.
    const top=Math.round(found[0].top),near=[...badges.keys()].find(value=>Math.abs(value-top)<16),entry=near===undefined?{threads:[],label:target.label}:badges.get(near)!;entry.threads.push(thread.id);badges.set(near??top,entry);
   }
   next.badges=[...badges].map(([top,entry])=>({top,...entry}));
   setLayer(previous=>JSON.stringify(previous)===JSON.stringify(next)?previous:next);
  });};
  const awareness=documentAwareness(doc),resize=new ResizeObserver(measure);resize.observe(element);
  doc.on('afterTransaction',measure);awareness?.on('change',measure);window.addEventListener('resize',measure);measure();
  return()=>{cancelAnimationFrame(frame);resize.disconnect();doc.off('afterTransaction',measure);awareness?.off('change',measure);window.removeEventListener('resize',measure);};
 // `peers`/`resolve` are read through the closure; `peerKey` re-arms only when a collaborator's position changes.
 },[paper,doc,peerKey]);
 const style=(value:Box)=>({left:value.left,top:value.top,width:value.width,height:value.height});
 // First child of the paper with zero size: pagination (which measures children) skips it, and the last block stays last.
 return <div className="word-collab-layer">
  {layer.comments.map((comment,i)=>comment.boxes.map((value,j)=><span key={'c'+i+':'+j} className="word-comment-mark" aria-hidden="true" style={style(value)}/>))}
  {layer.selections.map(selection=>selection.boxes.map((value,j)=><span key={'s'+selection.client+':'+j} className="word-peer-selection" aria-hidden="true" style={{...style(value),backgroundColor:selection.color}}/>))}
  {layer.carets.map(caret=><span key={'k'+caret.client} className="word-peer-caret" aria-hidden="true" style={{...style(caret.box),backgroundColor:caret.color}}><span className="word-peer-flag" style={{backgroundColor:caret.color}}>{caret.name}</span></span>)}
  {layer.badges.map(badge=><button key={'b'+badge.top} type="button" className="word-comment-badge" style={{top:badge.top,left:layer.margin}} aria-label={`查看批注：${badge.label}`+(badge.threads.length>1?`（${badge.threads.length} 条）`:'')} title={badge.label} onMouseDown={event=>{event.preventDefault();event.stopPropagation();}} onClick={event=>{event.stopPropagation();requestOfficeThread(workspace,id,badge.threads[0]);}}><MessageSquare size={13} aria-hidden="true"/>{badge.threads.length>1&&<span>{badge.threads.length}</span>}</button>)}
 </div>;
}
