import {Y} from '@mydoc/crdt';
import {textLinePosition,textLineStarts} from './office-text-lines';
export type TextPeer={client:number;name:string;color:string;anchor?:number;head?:number;startLine?:number;endLine?:number;label:string};
export function captureTextCursor(text:Y.Text,start:number,end:number,backward=false){
 const anchor=backward?end:start,head=backward?start:end;
 return {anchor:Y.createRelativePositionFromTypeIndex(text,anchor,backward&&end>start?-1:0),head:Y.createRelativePositionFromTypeIndex(text,head,!backward&&end>start?-1:0)};
}
export function textPeers(text:Y.Text,states:Map<number,Record<string,unknown>>,self:number){
 const value=text.toString(),starts=textLineStarts(value),peers:TextPeer[]=[];
 for(const[client,state]of states){
  if(client===self||!state?.user||typeof state.user!=='object')continue;
  const user=state.user as {name?:unknown;color?:unknown};if(typeof user.name!=='string')continue;
  const peer:TextPeer={client,name:user.name.slice(0,160),color:typeof user.color==='string'&&/^#[0-9a-f]{6}$/i.test(user.color)?user.color:'#3159cb',label:'正在查看'};
  if(state.cursor&&typeof state.cursor==='object'&&text.doc)try{
   const cursor=state.cursor as {anchor:unknown;head:unknown},anchor=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(cursor.anchor),text.doc),head=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(cursor.head),text.doc);
   if(anchor?.type===text&&head?.type===text){const from=textLinePosition(value,starts,Math.min(anchor.index,head.index)),to=textLinePosition(value,starts,Math.max(anchor.index,head.index));Object.assign(peer,{anchor:anchor.index,head:head.index,startLine:from.line,endLine:to.line,label:anchor.index===head.index?`第 ${from.line} 行，第 ${from.column} 列`:`选区：第 ${from.line} 行 ${from.column} 列至第 ${to.line} 行 ${to.column} 列`});}
  }catch{/* Stale or malformed remote anchors are displayed as viewers. */}
  peers.push(peer);if(peers.length===32)break;
 }
 return peers.sort((a,b)=>a.client-b.client);
}
