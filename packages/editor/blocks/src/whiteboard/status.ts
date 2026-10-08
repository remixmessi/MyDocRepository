import {groupChain} from './groups';
import type {SurfaceElement} from './model';
export interface GroupInfo {id:string;name:string;locked?:boolean;hidden?:boolean}
export const isHidden=(s:SurfaceElement)=>!!s.hidden||!!s.groupInfo?.some(g=>g.hidden);
export const isLocked=(s:SurfaceElement)=>!!s.locked||!!s.groupInfo?.some(g=>g.locked);
export function validStatus(s:SurfaceElement){return(s.layerName===undefined||typeof s.layerName==='string'&&s.layerName.length<=200)&&(s.locked===undefined||typeof s.locked==='boolean')&&(s.hidden===undefined||typeof s.hidden==='boolean')&&(s.groupInfo===undefined||Array.isArray(s.groupInfo)&&s.groupInfo.length<=16&&s.groupInfo.every(g=>!!g&&groupChain(s).includes(g.id)&&typeof g.name==='string'&&g.name.length<=200&&(g.locked===undefined||typeof g.locked==='boolean')&&(g.hidden===undefined||typeof g.hidden==='boolean'))&&new Set(s.groupInfo.map(g=>g.id)).size===s.groupInfo.length);}
export function setGroupInfo(elements:SurfaceElement[],id:string,patch:Partial<Omit<GroupInfo,'id'>>){return elements.filter(s=>groupChain(s).includes(id)).map(s=>({...s,groupInfo:groupChain(s).map(gid=>({id:gid,name:'',...s.groupInfo?.find(g=>g.id===gid),...(gid===id?patch:{})}))}));}
