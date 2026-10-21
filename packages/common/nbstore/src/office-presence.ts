/** Typed collaborator focus shared through presence. Positions use stable identities, never screen indexes. */
export type OfficeFocus=
 {kind:'word';text:string;anchor:unknown;head:unknown}|
 {kind:'cell';sheet:string;row:string;column:string;label:string;end?:{row:string;column:string;label:string}}|
 {kind:'slide';slide:string;object?:string}|
 {kind:'pdf';page:number};
const id=(value:unknown,max=180)=>typeof value==='string'&&/^[a-zA-Z0-9:_-]{1,180}$/.test(value)&&value.length<=max;
/** Shape check only; relative positions are validated separately by the caller's Yjs position check. */
export function validOfficeFocus(value:unknown,position:(value:unknown)=>boolean):value is OfficeFocus{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const v=value as Record<string,unknown>,keys=Object.keys(v);
 if(v.kind==='word')return keys.length===4&&id(v.text)&&position(v.anchor)&&position(v.head);
 if(v.kind==='cell'){const cell=(row:unknown,column:unknown,label:unknown)=>/^\d+-\d+$/.test(String(row))&&id(column)&&typeof label==='string'&&/^[A-Z]{1,3}\d{1,5}$/.test(label);
  // `end` is the opposite corner of a selected range, by stable row/column identity.
  const end=v.end as Record<string,unknown>|undefined,validEnd=v.end===undefined||!!end&&typeof end==='object'&&!Array.isArray(end)&&Object.keys(end).length===3&&cell(end.row,end.column,end.label);
  return keys.length===(v.end===undefined?5:6)&&keys.every(key=>['kind','sheet','row','column','label','end'].includes(key))&&id(v.sheet,128)&&cell(v.row,v.column,v.label)&&validEnd;}
 if(v.kind==='slide')return (keys.length===2||keys.length===3&&/^\d+-\d+$/.test(String(v.object)))&&id(v.slide,80)&&keys.every(key=>['kind','slide','object'].includes(key));
 return v.kind==='pdf'&&keys.length===2&&Number.isSafeInteger(v.page)&&Number(v.page)>=1&&Number(v.page)<=100000;
}
export type OfficePeer={client:number;name:string;color:string;focus?:OfficeFocus};
export function officePeers(states:Map<number,Record<string,unknown>>,self:number,kind:OfficeFocus['kind']){
 const peers:OfficePeer[]=[];
 for(const [client,state]of states){
  if(client===self||!state?.user||typeof state.user!=='object')continue;const user=state.user as {name?:unknown;color?:unknown};if(typeof user.name!=='string')continue;
  const focus=state.focus&&typeof state.focus==='object'&&(state.focus as OfficeFocus).kind===kind?state.focus as OfficeFocus:undefined;
  peers.push({client,name:user.name.slice(0,160),color:typeof user.color==='string'&&/^#[0-9a-f]{6}$/i.test(user.color)?user.color:'#3159cb',...(focus?{focus}:{})});if(peers.length===32)break;
 }
 return peers.sort((a,b)=>a.client-b.client);
}
