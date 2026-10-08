import {groupChain,withGroupChain,validGroupTree,normalizeGroups} from './groups';
import {isLocked} from './status';
import type {SurfaceElement} from './model';
export function transferGroup(ids:string[],elements:SurfaceElement[],target?:string){
 const selected=new Set(ids),members=elements.filter(s=>selected.has(s.id));if(!members.length)throw new Error('请先选择要移动的内容');if(members.some(isLocked))throw new Error('所选内容已锁定');if(members.some(s=>s.kind==='frame'))throw new Error('画框不参与分组');
 const targetMember=target?elements.find(s=>groupChain(s).includes(target)):undefined;if(target&&!targetMember)throw new Error('目标分组已不存在');const targetChain=targetMember?groupChain(targetMember).slice(0,groupChain(targetMember).indexOf(target!)+1):[];if(targetMember?.groupInfo?.some(g=>targetChain.includes(g.id)&&g.locked))throw new Error('目标分组已锁定');
 const complete=new Set<string>();for(const s of members)for(const g of groupChain(s))if(elements.filter(e=>groupChain(e).includes(g)).every(e=>selected.has(e.id)))complete.add(g);
 const changes=members.map(s=>{const chain=groupChain(s),index=chain.findIndex(g=>complete.has(g)),suffix=index<0?[]:chain.slice(index);if(targetChain.some(g=>suffix.includes(g)))throw new Error('不能把分组移入自身或子组');if(targetChain.length+suffix.length>16)throw new Error('分组最多支持 16 层');const next=withGroupChain(s,[...targetChain,...suffix]);const metadata=[...(targetMember?.groupInfo??[]).filter(g=>targetChain.includes(g.id)),...(s.groupInfo??[]).filter(g=>suffix.includes(g.id))];if(metadata.length)next.groupInfo=metadata.map(g=>({...g}));else delete next.groupInfo;return next;});
 const update=new Map(changes.map(s=>[s.id,s])),next=normalizeGroups(elements.map(s=>update.get(s.id)??s));if(!validGroupTree(next))throw new Error('移动后的分组层级无效');return next.filter(s=>s!==elements.find(e=>e.id===s.id));
}
