import {useEffect,useLayoutEffect,useRef,useState,type PointerEvent} from 'react';
import {MessageSquare} from 'lucide-react';
import {projectPreferences} from '@mydoc/nbstore';
import './styles.css';
const VISIBLE_KEY='mydoc:agent-launcher-visible:v1',POSITION_KEY='mydoc:agent-launcher-position:v1';
interface Position {x:number;y:number}
function readPosition():Position|undefined{try{const value=JSON.parse(projectPreferences.getItem(POSITION_KEY)??'null');if(value&&Number.isFinite(value.x)&&Number.isFinite(value.y))return {x:value.x,y:value.y};}catch{}return undefined;}
export function useAgentLauncherVisibility(){
 const read=()=>projectPreferences.getItem(VISIBLE_KEY)!=='false';
 const [visible,setVisible]=useState(read),[saving,setSaving]=useState(false),[error,setError]=useState('');
 useEffect(()=>{const sync=(event:Event)=>{if((event as CustomEvent<{key:string}>).detail?.key===VISIBLE_KEY)setVisible(read());};window.addEventListener('mydoc-preference',sync);return()=>window.removeEventListener('mydoc-preference',sync);},[]);
 const save=async(value:boolean)=>{if(saving)return;const previous=visible;setVisible(value);setSaving(true);setError('');try{await projectPreferences.setItem(VISIBLE_KEY,String(value));}catch{setVisible(previous);setError('图标显示设置未保存，请重试。');}finally{setSaving(false);}};
 return {visible,saving,error,save};
}
/** Remains mounted while chat is open, preserving its location across open/close. */
export function AgentLauncher({visible,onOpen}:{visible:boolean;onOpen:()=>void}){
 const [position,setPosition]=useState<Position|undefined>(readPosition),[error,setError]=useState('');
 const button=useRef<HTMLButtonElement>(null),suppressClick=useRef(false),last=useRef(position);
 const writes=useRef(Promise.resolve());
 const drag=useRef<{id:number;x:number;y:number;left:number;top:number;moved:boolean}>();
 const bound=(x:number,y:number):Position=>{const rect=button.current?.getBoundingClientRect();return {x:Math.max(8,Math.min(x,window.innerWidth-(rect?.width??110)-8)),y:Math.max(8,Math.min(y,window.innerHeight-(rect?.height??48)-8))};};
 const move=(value:Position)=>{last.current=value;setPosition(value);};
 const persist=(value:Position)=>{writes.current=writes.current.then(async()=>{try{await projectPreferences.setItem(POSITION_KEY,JSON.stringify(value));setError('');}catch{setError('图标位置未保存，刷新后将恢复原位置。');}});return writes.current;};
 useLayoutEffect(()=>{if(visible&&last.current)move(bound(last.current.x,last.current.y));},[visible]);
 useEffect(()=>{const resize=()=>{drag.current=undefined;if(last.current)move(bound(last.current.x,last.current.y));};window.addEventListener('resize',resize);return()=>window.removeEventListener('resize',resize);},[]);
 const start=(event:PointerEvent<HTMLButtonElement>)=>{if(event.button!==0||!event.isPrimary)return;const rect=event.currentTarget.getBoundingClientRect();drag.current={id:event.pointerId,x:event.clientX,y:event.clientY,left:rect.left,top:rect.top,moved:false};suppressClick.current=false;event.currentTarget.setPointerCapture(event.pointerId);event.currentTarget.focus();};
 const moving=(event:PointerEvent<HTMLButtonElement>)=>{const state=drag.current;if(!state||state.id!==event.pointerId)return;if(!state.moved&&Math.hypot(event.clientX-state.x,event.clientY-state.y)<6)return;state.moved=true;suppressClick.current=true;event.preventDefault();move(bound(state.left+event.clientX-state.x,state.top+event.clientY-state.y));};
 const end=(event:PointerEvent<HTMLButtonElement>,cancelled=false)=>{const state=drag.current;if(!state||state.id!==event.pointerId)return;drag.current=undefined;if(state.moved&&cancelled)move(bound(state.left,state.top));else if(state.moved&&last.current)void persist(last.current);};
 if(!visible)return null;
 return <><button ref={button} className="agent-launcher" style={position?{left:position.x,top:position.y,right:'auto',bottom:'auto'}:undefined} aria-label="Agent 助手" aria-expanded={false} aria-controls="mydoc-agent-window" aria-describedby="agent-launcher-help" onPointerDown={start} onPointerMove={moving} onPointerUp={event=>end(event)} onPointerCancel={event=>end(event,true)} onClick={event=>{if(suppressClick.current&&event.detail!==0){suppressClick.current=false;return;}onOpen();}} onKeyDown={event=>{if(!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key))return;event.preventDefault();const rect=event.currentTarget.getBoundingClientRect(),step=event.shiftKey?40:10;const value=bound(rect.left+(event.key==='ArrowLeft'?-step:event.key==='ArrowRight'?step:0),rect.top+(event.key==='ArrowUp'?-step:event.key==='ArrowDown'?step:0));move(value);void persist(value);}}><MessageSquare size={21}/><span>Agent</span></button><span id="agent-launcher-help" className="sr-only">点击打开助手；拖动或使用方向键移动位置。</span>{error&&<div className="agent-launcher-notice" role="alert">{error}<button onClick={()=>{if(last.current)void persist(last.current);}}>重试保存</button></div>}</>;
}
