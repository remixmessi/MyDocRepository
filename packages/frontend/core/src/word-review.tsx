import {useEffect,useState} from 'react';
import {Check,X} from 'lucide-react';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {cloudJson} from '../../../common/nbstore/src/cloud';
import {getStorageAccount} from '../../../common/nbstore/src/desktop';
import {resolveWordRevisions,wordRevisionColor,wordRevisions,type WordAuthor,type WordRevision} from '../../../common/nbstore/src/office-word-revision';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import type {LocalWorkspace} from './workspace';

/**
 * The author of tracked changes. Team workspaces use the signed-in account (the server checks new revisions
 * against it); personal workspaces have one local author.
 */
export function useWordAuthor(workspace:LocalWorkspace):WordAuthor|undefined{
 const team=!!workspace.sync,[author,setAuthor]=useState<WordAuthor|undefined>(()=>team?undefined:{id:'local',name:'本机用户'});
 useEffect(()=>{if(!team){setAuthor({id:'local',name:'本机用户'});return;}let alive=true;
  cloudJson<{user:{id:string;name:string}|null}>('/session').then(value=>{if(alive&&value.user)setAuthor({id:value.user.id,name:value.user.name||'协作者'});}).catch(()=>{const id=getStorageAccount();if(alive&&id)setAuthor({id,name:'协作者'});});
  return()=>{alive=false;};
 },[team,workspace]);
 return author;
}
export const setWordTracking=(doc:Y.Doc,on:boolean)=>doc.transact(()=>doc.getMap('pageSettings').set('trackChanges',on),LOCAL_ORIGIN);
/** Document order of a position, for previous/next navigation. */
function order(doc:Y.Doc){const texts=wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray());return (text:Y.Text,offset:number)=>texts.indexOf(text)*1e6+offset;}
/** The revision at the caret, or the next one after it (wrapping), like Word's 接受/拒绝 on the current change. */
export function revisionAt(doc:Y.Doc,caret:{text:Y.Text;offset:number}|undefined,direction:1|-1=1,strict=false):WordRevision|undefined{
 const list=wordRevisions(doc);if(!list.length)return;if(!caret)return direction>0?list[0]:list[list.length-1];const key=order(doc),at=key(caret.text,caret.offset);
 if(!strict){const inside=list.find(revision=>revision.text===caret.text&&caret.offset>=revision.start&&caret.offset<=revision.end);if(inside)return inside;}
 return direction>0?list.find(revision=>key(revision.text,revision.start)>at)??list[0]:[...list].reverse().find(revision=>key(revision.text,revision.end)<at)??list[list.length-1];
}
export function resolveRevision(doc:Y.Doc,accept:boolean,id?:string){return resolveWordRevisions(doc,accept,LOCAL_ORIGIN,id);}
/** 审阅窗格: every change with author, time and text; accept/reject one, or jump to it. */
export function WordReviewPane({doc,readOnly,onLocate,onClose}:{doc:Y.Doc;readOnly:boolean;onLocate:(revision:WordRevision)=>void;onClose:()=>void}){
 const list=wordRevisions(doc),groups=[...new Map(list.map(revision=>[revision.id,list.filter(item=>item.id===revision.id)])).values()];
 return <aside className="word-review-pane" aria-label="审阅窗格"><header><h3>修订 · {groups.length}</h3><button type="button" aria-label="关闭审阅窗格" onClick={onClose}><X size={14} aria-hidden="true"/></button></header>
  {!groups.length?<p>没有修订。打开“修订”后，插入和删除的文字会在这里列出。</p>:<ol>{groups.map(parts=>{const first=parts[0],kinds=new Set(parts.map(part=>part.kind));
   return <li key={first.id} style={{['--revision-color' as string]:wordRevisionColor(first.author)}}>
    <button type="button" className="word-review-item" onClick={()=>onLocate(first)}><strong>{first.name}</strong><span>{kinds.size>1?'替换':first.kind==='ins'?'插入':'删除'} · {new Date(first.date).toLocaleString()}</span>
     <span className="word-review-text">{parts.map((part,i)=><span key={i} data-revision={part.kind}>{part.excerpt||'（空）'}</span>)}</span></button>
    <span className="word-review-actions"><button type="button" aria-label={`接受 ${first.name} 的修订`} title="接受" disabled={readOnly} onClick={()=>resolveRevision(doc,true,first.id)}><Check size={14} aria-hidden="true"/></button><button type="button" aria-label={`拒绝 ${first.name} 的修订`} title="拒绝" disabled={readOnly} onClick={()=>resolveRevision(doc,false,first.id)}><X size={14} aria-hidden="true"/></button></span>
   </li>;})}</ol>}
 </aside>;
}
