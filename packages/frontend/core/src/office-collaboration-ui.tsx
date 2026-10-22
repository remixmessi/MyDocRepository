import {useEffect,useRef,useState} from 'react';
import {createPortal} from 'react-dom';
import {Bell,History} from 'lucide-react';
import type {Y} from '@mydoc/crdt';
import {cloudJson} from '../../../common/nbstore/src/cloud';
import type {OfficePeer} from '../../../common/nbstore/src/office-presence';
import type {LocalWorkspace} from './workspace';

export type OfficeActivity={user:string;userName:string;started:number;last:number;edits:number};
export type CloudNotification={id:string;workspace:string;workspaceName:string;resource:string;resourceName:string;kind:'mention'|'reply';actor:string;actorName:string;thread:string;excerpt:string;created:number;read:boolean};

/** A notification click may arrive before the target editor mounts, so the request waits here until claimed. */
let pendingThread:{workspace:string;id:string;thread:string}|undefined;
export function requestOfficeThread(workspace:string,id:string,thread:string){pendingThread={workspace,id,thread};window.dispatchEvent(new CustomEvent('office-thread-request'));}
export function claimOfficeThread(workspace:string,id:string){const request=pendingThread;if(request?.workspace!==workspace||request.id!==id)return;pendingThread=undefined;return request.thread;}

const when=(at:number)=>{const minutes=Math.round((Date.now()-at)/60000);if(minutes<1)return '刚刚';if(minutes<60)return minutes+' 分钟前';const hours=Math.round(minutes/60);return hours<24?hours+' 小时前':new Date(at).toLocaleString();};

/** Recent edit sessions for a team office file. Reloads after remote or local document updates, at most every few seconds. */
export function useOfficeActivity(doc:Y.Doc,workspace:LocalWorkspace,id:string,open=false){
 const [items,setItems]=useState<OfficeActivity[]>([]),[error,setError]=useState(''),enabled=!!workspace.sync,reload=useRef<()=>void>(()=>{});
 useEffect(()=>{if(!enabled)return;let alive=true,timer:ReturnType<typeof setTimeout>|undefined,last=0,failures=0;
  const schedule=(delay:number)=>{if(!timer)timer=setTimeout(()=>{timer=undefined;load();},delay);};
  // A failed read (e.g. while the file is still opening) retries with backoff instead of waiting for the next edit.
  const load=()=>{last=Date.now();cloudJson<OfficeActivity[]>('/office-activity?'+new URLSearchParams({workspace:workspace.id,id})).then(value=>{if(!alive)return;failures=0;setItems(value);setError('');}).catch(reason=>{if(!alive)return;setError(reason instanceof Error?reason.message:'无法读取编辑记录');if(failures<5)schedule(2000*2**failures++);});};
  const changed=()=>schedule(Math.max(1500,5000-(Date.now()-last)));
  reload.current=()=>{clearTimeout(timer);timer=undefined;load();};
  load();doc.on('update',changed);const idle=setInterval(changed,60000);return()=>{alive=false;clearTimeout(timer);clearInterval(idle);doc.off('update',changed);};
 },[doc,workspace,id,enabled]);
 useEffect(()=>{if(open)reload.current();},[open]);
 return {items,error,enabled};
}

/** "最近编辑" button in the title bar with a popover of recent edit sessions. Team workspaces only. */
export function OfficeActivityButton({doc,workspace,id}:{doc:Y.Doc;workspace:LocalWorkspace;id:string}){
 const [open,setOpen]=useState(false),{items,error,enabled}=useOfficeActivity(doc,workspace,id,open),root=useRef<HTMLDivElement>(null);
 useEffect(()=>{if(!open)return;const outside=(event:MouseEvent)=>{if(!root.current?.contains(event.target as Node))setOpen(false);},key=(event:KeyboardEvent)=>{if(event.key==='Escape')setOpen(false);};document.addEventListener('mousedown',outside);document.addEventListener('keydown',key);return()=>{document.removeEventListener('mousedown',outside);document.removeEventListener('keydown',key);};},[open]);
 if(!enabled)return null;const latest=items[0];
 return <div className="office-activity" ref={root}><button type="button" aria-expanded={open} aria-label={latest?`最近编辑：${latest.userName}，${when(latest.last)}`:'编辑记录'} title="编辑记录" onClick={()=>setOpen(value=>!value)}><History size={15} aria-hidden="true"/><span>{latest?`${latest.userName} · ${when(latest.last)}`:'编辑记录'}</span></button>
  {open&&<div className="office-activity-popover" role="dialog" aria-label="编辑记录"><h3>编辑记录</h3>{error&&<p role="alert">{error}</p>}{!items.length&&!error&&<p>暂无团队成员的编辑记录。</p>}<ol>{items.map(item=><li key={item.user+':'+item.started}><strong>{item.userName}</strong><span>{when(item.last)}</span><small>{new Date(item.started).toLocaleTimeString()}–{new Date(item.last).toLocaleTimeString()} · {item.edits} 次保存</small></li>)}</ol><p className="office-activity-note">相邻 10 分钟内的连续编辑合并为一条，每个文件保留最近 200 条。</p></div>}
 </div>;
}

/**
 * Follow a collaborator: whenever their focus changes, jump to it. Stops when they leave or when the user
 * presses a key or clicks inside the editor, so following never fights the user's own editing.
 */
export function useFollowPeer(peers:OfficePeer[],locate:((peer:OfficePeer)=>void)|undefined,root:{current:HTMLElement|null}){
 const [following,setFollowing]=useState<number>(),peer=peers.find(value=>value.client===following),focus=peer?.focus?JSON.stringify(peer.focus):'';
 useEffect(()=>{if(following!==undefined&&!peer)setFollowing(undefined);},[following,peer]);
 // Only a focus change (or starting to follow) jumps; re-renders from unrelated peers do not.
 useEffect(()=>{if(peer&&focus)locate?.(peer);},[focus,following]);
 useEffect(()=>{const element=root.current;if(following===undefined||!element)return;const stop=(event:Event)=>{if(event.isTrusted&&!(event.target as HTMLElement).closest?.('.office-presence-list'))setFollowing(undefined);};element.addEventListener('keydown',stop,true);element.addEventListener('mousedown',stop,true);return()=>{element.removeEventListener('keydown',stop,true);element.removeEventListener('mousedown',stop,true);};},[following,root]);
 return {following,follow:(client:number|undefined)=>setFollowing(client)};
}

/** Notification bell for @mentions and replies in team office files. Hidden until a team workspace exists. */
export function NotificationBell({enabled,onOpen}:{enabled:boolean;onOpen:(notification:CloudNotification)=>void}){
 const [state,setState]=useState<{items:CloudNotification[];unread:number}>({items:[],unread:0}),[open,setOpen]=useState(false),[error,setError]=useState(''),[place,setPlace]=useState<{top:number;left:number}>(),root=useRef<HTMLDivElement>(null),popover=useRef<HTMLDivElement>(null),button=useRef<HTMLButtonElement>(null);
 const load=useRef<()=>void>(()=>{});
 useEffect(()=>{if(!enabled){setState({items:[],unread:0});return;}let alive=true;
  load.current=()=>{cloudJson<{items:CloudNotification[];unread:number}>('/notifications?limit=50').then(value=>{if(alive){setState(value);setError('');}}).catch(reason=>{if(alive)setError(reason instanceof Error?reason.message:'无法读取通知');});};
  // Every open team document has its own event stream, so one notification can arrive several times at once.
  let pending:ReturnType<typeof setTimeout>|undefined;const changed=()=>{clearTimeout(pending);pending=setTimeout(()=>load.current(),300);};
  load.current();window.addEventListener('cloud-notifications-changed',changed);const timer=setInterval(changed,60000);
  return()=>{alive=false;clearTimeout(pending);window.removeEventListener('cloud-notifications-changed',changed);clearInterval(timer);};
 },[enabled]);
 // The sidebar clips overflow, so the panel is portalled to the body and placed from the bell's position.
 useEffect(()=>{if(!open)return;const rect=button.current?.getBoundingClientRect();if(rect)setPlace({top:rect.bottom+6,left:Math.max(8,Math.min(rect.left,window.innerWidth-348))});load.current();const outside=(event:MouseEvent)=>{if(!root.current?.contains(event.target as Node)&&!popover.current?.contains(event.target as Node))setOpen(false);},key=(event:KeyboardEvent)=>{if(event.key==='Escape'){setOpen(false);button.current?.focus();}};document.addEventListener('mousedown',outside);document.addEventListener('keydown',key);return()=>{document.removeEventListener('mousedown',outside);document.removeEventListener('keydown',key);};},[open]);
 if(!enabled)return null;
 const read=async(ids?:string[])=>{try{await cloudJson('/notifications-read','POST',ids?{ids}:{});load.current();}catch(reason){setError(reason instanceof Error?reason.message:'无法更新通知');}};
 return <div className="notification-bell" ref={root}><button ref={button} type="button" className="notification-bell-button" aria-label={state.unread?`通知，${state.unread} 条未读`:'通知'} aria-expanded={open} onClick={()=>setOpen(value=>!value)}><Bell size={16} aria-hidden="true"/>{state.unread>0&&<span className="notification-count" aria-hidden="true">{state.unread>99?'99+':state.unread}</span>}</button>
  {open&&createPortal(<div ref={popover} className="notification-popover" role="dialog" aria-label="通知" style={place}><div className="notification-heading"><h3>通知</h3><button type="button" disabled={!state.unread} onClick={()=>void read()}>全部标为已读</button></div>{error&&<p role="alert">{error}</p>}
   {!state.items.length&&!error?<p className="notification-empty">暂无通知。在办公文件评论中被 @ 提及或有人回复你参与的讨论时会显示在这里。</p>:<ul>{state.items.map(item=><li key={item.id} data-read={item.read||undefined}><button type="button" onClick={()=>{setOpen(false);if(!item.read)void read([item.id]);onOpen(item);}}><span className="notification-title"><strong>{item.actorName}</strong>{item.kind==='mention'?' 在评论中提到了你':' 回复了讨论'}</span><span className="notification-excerpt">{item.excerpt}</span><span className="notification-meta">{item.workspaceName} · {item.resourceName} · {when(item.created)}</span></button></li>)}</ul>}
  </div>,document.body)}
 </div>;
}
