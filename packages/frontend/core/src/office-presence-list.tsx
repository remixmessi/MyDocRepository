import {Eye,EyeOff} from 'lucide-react';
import type {OfficePeer} from '../../../common/nbstore/src/office-presence';
const initials=(name:string)=>{const value=name.trim();if(!value)return '?';return /^[\x00-\x7f]/.test(value)?value.split(/\s+/).slice(0,2).map(part=>part[0]).join('').toUpperCase():value.slice(-2);};
/** Online collaborators of this file; clicking a located peer jumps to their position, the eye button follows them. */
export function OfficePresenceList({peers,describe,onLocate,following,onFollow}:{peers:OfficePeer[];describe:(peer:OfficePeer)=>string|undefined;onLocate?:(peer:OfficePeer)=>void;following?:number;onFollow?:(client:number|undefined)=>void}){
 if(!peers.length)return null;
 return <div className="text-presence-list office-presence-list" role="group" aria-label="在线协作成员"><span>其他在线会话 · {peers.length}{peers.length===32?'（最多显示 32 个）':''}</span>{peers.map(peer=>{const label=describe(peer),followed=following===peer.client;return <span key={peer.client} className="office-peer" data-following={followed||undefined} style={{['--peer-color' as string]:peer.color}}>
  <button title={peer.name+' · '+(label??'正在查看')} disabled={!label||!onLocate} onClick={()=>onLocate?.(peer)}><span className="office-peer-avatar" aria-hidden="true" style={{backgroundColor:peer.color}}>{initials(peer.name)}</span><span>{peer.name} · {label??'正在查看'}</span></button>
  {onFollow&&<button className="office-peer-follow" aria-pressed={followed} aria-label={(followed?'停止跟随 ':'跟随 ')+peer.name} title={followed?'停止跟随':'跟随此成员的位置'} disabled={!label&&!followed} onClick={()=>onFollow(followed?undefined:peer.client)}>{followed?<EyeOff size={13} aria-hidden="true"/>:<Eye size={13} aria-hidden="true"/>}</button>}
 </span>;})}</div>;
}
