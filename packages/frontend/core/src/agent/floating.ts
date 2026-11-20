import {useEffect,useRef,useState,type PointerEvent,type KeyboardEvent} from 'react';
export function useFloatingAgent(){
 const [position,setPosition]=useState<{x:number;y:number}>();
 const element=useRef<HTMLElement>(null),drag=useRef<{id:number;x:number;y:number;left:number;top:number}>();
 const clamp=(x:number,y:number)=>{const rect=element.current?.getBoundingClientRect();return {x:Math.max(8,Math.min(x,window.innerWidth-(rect?.width??420)-8)),y:Math.max(8,Math.min(y,window.innerHeight-(rect?.height??600)-8))};};
 useEffect(()=>{const resize=()=>{drag.current=undefined;setPosition(old=>old?clamp(old.x,old.y):undefined);};window.addEventListener('resize',resize);return()=>window.removeEventListener('resize',resize);},[]);
 const onPointerDown=(event:PointerEvent<HTMLElement>)=>{if(event.button!==0||window.innerWidth<=640||(event.target as HTMLElement).closest('button'))return;const rect=element.current!.getBoundingClientRect();drag.current={id:event.pointerId,x:event.clientX,y:event.clientY,left:rect.left,top:rect.top};event.currentTarget.setPointerCapture(event.pointerId);event.preventDefault();};
 const onPointerMove=(event:PointerEvent<HTMLElement>)=>{const start=drag.current;if(start?.id===event.pointerId)setPosition(clamp(start.left+event.clientX-start.x,start.top+event.clientY-start.y));};
 const onKeyDown=(event:KeyboardEvent<HTMLElement>)=>{if(!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)||window.innerWidth<=640)return;event.preventDefault();const rect=element.current!.getBoundingClientRect(),step=event.shiftKey?40:10;setPosition(clamp(rect.left+(event.key==='ArrowLeft'?-step:event.key==='ArrowRight'?step:0),rect.top+(event.key==='ArrowUp'?-step:event.key==='ArrowDown'?step:0)));};
 return {element,style:position?{left:position.x,top:position.y,right:'auto',bottom:'auto'}:undefined,handle:{onPointerDown,onPointerMove,onPointerUp:()=>{drag.current=undefined;},onPointerCancel:()=>{drag.current=undefined;},onKeyDown}};
}
