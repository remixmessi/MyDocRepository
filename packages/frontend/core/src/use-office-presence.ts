import {useEffect,useReducer} from 'react';
import {documentAwareness,type Y} from '@mydoc/crdt';
import {officePeers,type OfficeFocus} from '../../../common/nbstore/src/office-presence';
/** Publish this view's typed focus and return other collaborators' focus of the same kind. No-op without team sync. */
export function useOfficePresence(doc:Y.Doc,kind:OfficeFocus['kind'],focus:OfficeFocus|undefined,enabled=true){
 const [,refresh]=useReducer(n=>n+1,0),awareness=documentAwareness(doc),serialized=focus?JSON.stringify(focus):'';
 useEffect(()=>{if(!awareness||!enabled)return;const changed=({added,updated,removed}:{added:number[];updated:number[];removed:number[]})=>{if([...added,...updated,...removed].some(client=>client!==doc.clientID))refresh();};awareness.on('change',changed);const hidden=()=>{if(document.visibilityState==='hidden'&&awareness.getLocalState())awareness.setLocalStateField('focus',null);};document.addEventListener('visibilitychange',hidden);refresh();return()=>{awareness.off('change',changed);document.removeEventListener('visibilitychange',hidden);if(awareness.getLocalState())awareness.setLocalStateField('focus',null);};},[awareness,enabled,doc]);
 useEffect(()=>{if(!awareness||!enabled||document.visibilityState==='hidden')return;const timer=setTimeout(()=>{if(awareness.getLocalState())awareness.setLocalStateField('focus',serialized?JSON.parse(serialized):null);},120);return()=>clearTimeout(timer);},[awareness,enabled,serialized]);
 return enabled&&awareness?officePeers(awareness.getStates() as Map<number,Record<string,unknown>>,doc.clientID,kind):[];
}
