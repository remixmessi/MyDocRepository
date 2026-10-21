import {useEffect,useReducer,useState,type RefObject} from 'react';
import {documentAwareness,type Y} from '@mydoc/crdt';
import {captureTextCursor,textPeers} from './office-text-presence';

const views=new WeakMap<Y.Doc,Set<symbol>>();
export function useTextPresence(text:Y.Text,input:RefObject<HTMLTextAreaElement>,selection:{start:number;end:number;direction:string},composing:boolean,enabled:boolean){
 const [focused,setFocused]=useState(false);
 const [,refresh]=useReducer(n=>n+1,0),awareness=text.doc&&documentAwareness(text.doc);
 useEffect(()=>{const element=input.current;if(!element)return;const focus=()=>setFocused(true),blur=()=>setFocused(false);element.addEventListener('focus',focus);element.addEventListener('blur',blur);setFocused(document.activeElement===element);return()=>{element.removeEventListener('focus',focus);element.removeEventListener('blur',blur);};},[input]);
 useEffect(()=>{
  if(!awareness||!text.doc||!enabled)return;
  const doc=text.doc,token=Symbol(),owners=views.get(doc)??new Set<symbol>();owners.add(token);views.set(doc,owners);
  const visible=()=>{if(document.visibilityState==='hidden'){if(awareness.getLocalState())awareness.setLocalState(null);}else if(!awareness.getLocalState()){const element=input.current;awareness.setLocalState({user:{name:'我',color:'#3159cb'},cursor:element&&document.activeElement===element&&element.value===text.toString()?captureTextCursor(text,element.selectionStart,element.selectionEnd,element.selectionDirection==='backward'):null});}};
  awareness.on('change',refresh);document.addEventListener('visibilitychange',visible);visible();refresh();
  return()=>{awareness.off('change',refresh);document.removeEventListener('visibilitychange',visible);owners.delete(token);if(!owners.size){awareness.setLocalState(null);views.delete(doc);}};
 },[text,input,awareness,enabled]);
 useEffect(()=>{
  if(!awareness||!enabled||composing||input.current!==document.activeElement||document.visibilityState==='hidden')return;
  const cursor=captureTextCursor(text,selection.start,selection.end,selection.direction==='backward');
  const timer=setTimeout(()=>{if(input.current===document.activeElement&&awareness.getLocalState())awareness.setLocalStateField('cursor',cursor);},120);
  return()=>clearTimeout(timer);
 },[text,input,awareness,selection.start,selection.end,selection.direction,composing,enabled,focused]);
 return enabled&&awareness?textPeers(text,awareness.getStates(),text.doc!.clientID):[];
}
