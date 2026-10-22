import type {TextPeer} from './office-text-presence';
import {useEffect,useLayoutEffect,useState,type RefObject} from 'react';
import {visibleTextLines} from './office-text-lines';

export function TextLineGutter({input,count,value,peers=[]}:{input:RefObject<HTMLTextAreaElement>;count:number;value:string;peers?:TextPeer[]}){
 const [view,setView]=useState({top:0,height:320,lineHeight:25.5,padding:24});
 useLayoutEffect(()=>{
  const element=input.current;if(!element)return;
  const style=getComputedStyle(element);setView({top:element.scrollTop,height:element.clientHeight,lineHeight:parseFloat(style.lineHeight)||25.5,padding:parseFloat(style.paddingTop)||0});
 },[input,value]);
 useEffect(()=>{
  const element=input.current;if(!element)return;let frame:number|undefined;
  const read=()=>{frame=undefined;const style=getComputedStyle(element);setView({top:element.scrollTop,height:element.clientHeight,lineHeight:parseFloat(style.lineHeight)||25.5,padding:parseFloat(style.paddingTop)||0});};
  const schedule=()=>{if(frame===undefined)frame=requestAnimationFrame(read);};
  const resize=new ResizeObserver(schedule);resize.observe(element);element.addEventListener('scroll',schedule,{passive:true});schedule();
  return()=>{resize.disconnect();element.removeEventListener('scroll',schedule);if(frame!==undefined)cancelAnimationFrame(frame);};
 },[input]);
 const {first,last}=visibleTextLines(count,view.top,view.height,view.lineHeight,view.padding);
 return <div className="text-line-gutter" aria-hidden="true" style={{width:`${Math.max(2,String(count).length)+2}ch`}}>{Array.from({length:Math.max(0,last-first)},(_,i)=>{
  const index=first+i;return <span key={index} data-line={index+1} style={{top:view.padding+index*view.lineHeight-view.top,height:view.lineHeight}}>{peers.filter(peer=>peer.startLine!==undefined&&index+1>=peer.startLine&&index+1<=peer.endLine!).slice(0,3).map((peer,marker)=><i key={peer.client} className="text-peer-line" title={peer.name+' · '+peer.label} style={{backgroundColor:peer.color,left:marker*4}}/>)}{index+1}</span>;
 })}</div>;
}
