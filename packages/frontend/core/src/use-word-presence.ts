import {useEffect,useReducer,type RefObject} from 'react';
import {Y,documentAwareness} from '@mydoc/crdt';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import {officePeers,type OfficePeer} from '../../../common/nbstore/src/office-presence';
import {wordTextIdentity} from './office-word-flow';
import {wordDomSelection} from './word-rich-input';
/**
 * Publish the caret as relative positions in the focused paragraph or cell, straight to awareness so typing and
 * selecting never re-render the editor; only other collaborators' changes trigger a render.
 */
export function useWordPresence(root:RefObject<HTMLElement>,doc:Y.Doc,enabled:boolean){
 const [,refresh]=useReducer(n=>n+1,0),awareness=enabled?documentAwareness(doc):undefined;
 useEffect(()=>{if(!awareness)return;let timer:ReturnType<typeof setTimeout>|undefined,last='';
  const publish=()=>{clearTimeout(timer);timer=setTimeout(()=>{if(!awareness.getLocalState()||document.visibilityState==='hidden')return;const element=root.current,dom=window.getSelection();let focus:unknown=null;
   const input=dom?.rangeCount?(dom.anchorNode instanceof HTMLElement?dom.anchorNode:dom.anchorNode?.parentElement)?.closest<HTMLElement>('.word-rich-input'):null;
   if(element&&input&&element.contains(input)){const text=wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray()).find(value=>wordTextIdentity(value)===input.dataset.wordText),range=wordDomSelection(input);if(text&&range)focus={kind:'word',text:input.dataset.wordText!,anchor:Y.relativePositionToJSON(Y.createRelativePositionFromTypeIndex(text,range.start)),head:Y.relativePositionToJSON(Y.createRelativePositionFromTypeIndex(text,range.end,range.end>range.start?-1:0))};}
   const serialized=JSON.stringify(focus);if(serialized!==last){last=serialized;awareness.setLocalStateField('focus',focus);}},120);};
  const changed=({added,updated,removed}:{added:number[];updated:number[];removed:number[]})=>{if([...added,...updated,...removed].some(client=>client!==doc.clientID))refresh();};
  document.addEventListener('selectionchange',publish);awareness.on('change',changed);refresh();
  return()=>{clearTimeout(timer);document.removeEventListener('selectionchange',publish);awareness.off('change',changed);if(awareness.getLocalState())awareness.setLocalStateField('focus',null);};
 },[root,doc,awareness]);
 const peers=awareness?officePeers(awareness.getStates() as Map<number,Record<string,unknown>>,doc.clientID,'word'):[];
 const resolve=(peer:OfficePeer)=>{if(peer.focus?.kind!=='word')return;try{const anchor=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(peer.focus.anchor),doc),head=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(peer.focus.head),doc);if(!anchor||anchor.type!==head?.type||!(anchor.type instanceof Y.Text))return;const texts=wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray()),index=texts.indexOf(anchor.type);return index<0?undefined:{text:anchor.type,index,start:Math.min(anchor.index,head.index),end:Math.max(anchor.index,head.index)};}catch{return;}};
 return {peers,resolve};
}
