import {useEffect,useState} from 'react';
import {Check,X} from 'lucide-react';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {cloudJson} from '../../../common/nbstore/src/cloud';
import {getStorageAccount} from '../../../common/nbstore/src/desktop';
import {resolveWordRevisions,wordRevisionColor,wordRevisions,wordRevisionOrder,type WordAuthor,type WordRevision,type WordRevisionKind} from '../../../common/nbstore/src/office-word-revision';
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
/** The revision at the caret, or the next one after it (wrapping), like Word's 接受/拒绝 on the current change. */
export function revisionAt(doc:Y.Doc,caret:{text:Y.Text;offset:number}|undefined,direction:1|-1=1,strict=false):WordRevision|undefined{
 const list=wordRevisions(doc);if(!list.length)return;if(!caret)return direction>0?list[0]:list[list.length-1];const at=wordRevisionOrder(doc)(caret.text,caret.offset);
 if(!strict){const inside=list.find(revision=>revision.text===caret.text&&caret.offset>=revision.start&&caret.offset<=revision.end);if(inside)return inside;}
 return direction>0?list.find(revision=>revision.order>at)??list[0]:[...list].reverse().find(revision=>revision.order+(revision.end-revision.start)<at)??list[list.length-1];
}
const KIND_LABEL:Record<WordRevisionKind,string>={ins:'插入',del:'删除',fmt:'格式',pfmt:'段落格式',insBreak:'拆分段落',delBreak:'合并段落',insImage:'插入图片',delImage:'删除图片'};
const FORMAT_LABEL:Record<string,string>={bold:'加粗',italic:'倾斜',underline:'下划线',strike:'删除线',fontSize:'字号',color:'颜色',highlight:'突出显示',vertAlign:'上下标',fontFamily:'字体',heading:'样式',align:'对齐',leftIndent:'左缩进',rightIndent:'右缩进',firstLineIndent:'特殊缩进',list:'列表',lineSpacing:'行距',spaceBefore:'段前',spaceAfter:'段后',pageBreakBefore:'段前分页'};
/** "格式：加粗、字号" — what a formatting revision changed. */
export function revisionSummary(parts:WordRevision[]){
 const kinds=new Set(parts.map(part=>part.kind));if(kinds.size===2&&kinds.has('ins')&&kinds.has('del'))return '替换';
 return [...kinds].map(kind=>{const keys=[...new Set(parts.filter(part=>part.kind===kind).flatMap(part=>Object.keys(part.old??{})))];return KIND_LABEL[kind]+(keys.length?'：'+keys.map(key=>FORMAT_LABEL[key]??key).join('、'):'');}).join('，');
}
export function resolveRevision(doc:Y.Doc,accept:boolean,id?:string){return resolveWordRevisions(doc,accept,LOCAL_ORIGIN,id);}
/** 审阅窗格: every change with author, time and text; accept/reject one, or jump to it. */
export function WordReviewPane({doc,readOnly,onLocate,onClose}:{doc:Y.Doc;readOnly:boolean;onLocate:(revision:WordRevision)=>void;onClose:()=>void}){
 const list=wordRevisions(doc),groups=[...new Map(list.map(revision=>[revision.id,list.filter(item=>item.id===revision.id)])).values()];
 return <aside className="word-review-pane" aria-label="审阅窗格"><header><h3>修订 · {groups.length}</h3><button type="button" aria-label="关闭审阅窗格" onClick={onClose}><X size={14} aria-hidden="true"/></button></header>
  {!groups.length?<p>没有修订。打开“修订”后，插入和删除的文字、格式更改、分段合段及图片的插入删除会在这里列出。</p>:<ol>{groups.map(parts=>{const first=parts[0];
   return <li key={first.id} style={{['--revision-color' as string]:wordRevisionColor(first.author)}}>
    <button type="button" className="word-review-item" onClick={()=>onLocate(first)}><strong>{first.name}</strong><span>{revisionSummary(parts)} · {new Date(first.date).toLocaleString()}</span>
     <span className="word-review-text">{parts.map((part,i)=><span key={i} data-revision={part.kind}>{part.excerpt||(part.kind==='insBreak'||part.kind==='delBreak'?'¶':'（空）')}</span>)}</span></button>
    <span className="word-review-actions"><button type="button" aria-label={`接受 ${first.name} 的修订`} title="接受" disabled={readOnly} onClick={()=>resolveRevision(doc,true,first.id)}><Check size={14} aria-hidden="true"/></button><button type="button" aria-label={`拒绝 ${first.name} 的修订`} title="拒绝" disabled={readOnly} onClick={()=>resolveRevision(doc,false,first.id)}><X size={14} aria-hidden="true"/></button></span>
   </li>;})}</ol>}
 </aside>;
}
