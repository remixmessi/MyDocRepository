import type {RefObject} from 'react';
import type {TextPeer} from './office-text-presence';
export function TextPresenceList({peers,input}:{peers:TextPeer[];input:RefObject<HTMLTextAreaElement>}){
 if(!peers.length)return null;
 return <div className="text-presence-list" role="group" aria-label="文本协作成员"><span>其他在线会话 · {peers.length}{peers.length===32?'（最多显示 32 个）':''}</span>{peers.map(peer=><button key={peer.client} title={peer.name+' · '+peer.label} disabled={peer.anchor===undefined} onClick={()=>{
  if(peer.anchor===undefined||peer.head===undefined)return;const element=input.current;if(!element)return;element.focus();element.setSelectionRange(Math.min(peer.anchor,peer.head),Math.max(peer.anchor,peer.head),peer.anchor>peer.head?'backward':'forward');
  const style=getComputedStyle(element);element.scrollTop=Math.max(0,((peer.startLine??1)-1)*(parseFloat(style.lineHeight)||25.5)-(element.clientHeight/3));
 }}><span className="text-peer-dot" aria-hidden="true" style={{backgroundColor:peer.color}}/><span>{peer.name} · {peer.label}</span></button>)}</div>;
}
