import type {SurfaceElement} from './model';
// Outer ancestors are optional; groupId remains the immediate parent for legacy data.
export const groupChain=(s:SurfaceElement):string[]=>s.groupId?[...(s.groupPath??[]),s.groupId]:[];
export function withGroupChain(s:SurfaceElement,chain:string[]):SurfaceElement {
  const copy={...s};delete copy.groupId;delete copy.groupPath;if(copy.groupInfo){copy.groupInfo=copy.groupInfo.filter(g=>chain.includes(g.id));if(!copy.groupInfo.length)delete copy.groupInfo;}
  if(chain.length){copy.groupId=chain.at(-1);if(chain.length>1)copy.groupPath=chain.slice(0,-1);}
  return copy;
}
export function validGroupPath(s:SurfaceElement){
  if(s.groupPath===undefined)return true;
  return !!s.groupId&&Array.isArray(s.groupPath)&&s.groupPath.length>0&&s.groupPath.length<16&&s.groupPath.every(id=>typeof id==='string'&&!!id&&id.length<=200)&&new Set(groupChain(s)).size===s.groupPath.length+1;
}
export function validGroupTree(elements:SurfaceElement[],allowSingletons=false){
  const info=new Map<string,string>();for(const s of elements)for(const id of groupChain(s)){const g=s.groupInfo?.find(g=>g.id===id),signature=JSON.stringify([g?.name??'',g?.locked??false,g?.hidden??false]);if(info.has(id)&&info.get(id)!==signature)return false;info.set(id,signature);}
  const parents=new Map<string,string|undefined>(),children=new Map<string,Set<string>>(),elementIds=new Set(elements.map(s=>s.id));
  for(const s of elements){if(!validGroupPath(s)||s.kind==='frame'&&s.groupId)return false;const chain=groupChain(s);
    for(let i=0;i<chain.length;i++){const id=chain[i],parent=chain[i-1];if(elementIds.has(id)||parents.has(id)&&parents.get(id)!==parent)return false;parents.set(id,parent);if(!children.has(id))children.set(id,new Set());children.get(id)!.add(s.id);}
  }
  return allowSingletons||[...children.values()].every(set=>set.size>=2);
}
// Collapse groups with fewer than two leaf members; a parent containing one subgroup is meaningful.
export function normalizeGroups(elements:SurfaceElement[]){
  const children=new Map<string,Set<string>>();for(const s of elements){const chain=groupChain(s);chain.forEach((id,i)=>{if(!children.has(id))children.set(id,new Set());children.get(id)!.add(s.id);});}
  const redundant=new Set([...children].filter(([,set])=>set.size<2).map(([id])=>id));
  return elements.map(s=>groupChain(s).some(id=>redundant.has(id))?withGroupChain(s,groupChain(s).filter(id=>!redundant.has(id))):s);
}
export const scopeElements=(elements:SurfaceElement[],scope?:string)=>scope?elements.filter(s=>groupChain(s).includes(scope)):elements;
export function selectionUnit(s:SurfaceElement,scope?:string){const chain=groupChain(s);if(scope&&!chain.includes(scope))return undefined;return chain[scope?chain.indexOf(scope)+1:0];}
export function expandGroups(ids:string[],elements:SurfaceElement[],scope?:string){
  const eligible=scopeElements(elements,scope),set=new Set(ids),groups=new Set(eligible.filter(s=>set.has(s.id)).map(s=>selectionUnit(s,scope)).filter(Boolean));
  return eligible.filter(s=>set.has(s.id)||groups.has(selectionUnit(s,scope))).map(s=>s.id);
}
export function selectedGroup(ids:string[],elements:SurfaceElement[],scope?:string){
  const set=new Set(ids),selected=elements.filter(s=>set.has(s.id));if(!selected.length)return undefined;const group=selectionUnit(selected[0],scope);
  return group&&selected.every(s=>selectionUnit(s,scope)===group)&&elements.filter(s=>groupChain(s).includes(group)).every(s=>set.has(s.id))?group:undefined;
}
export function placeInScope(s:SurfaceElement,elements:SurfaceElement[],scope?:string){
  if(!scope)return s;const member=elements.find(s=>groupChain(s).includes(scope));if(!member)throw new Error('当前分组已变化，请重新选择');const chain=groupChain(member);const ancestors=chain.slice(0,chain.indexOf(scope)+1);if(member.groupInfo?.some(g=>ancestors.includes(g.id)&&g.locked))throw new Error('当前分组已锁定，请先解锁');const placed=withGroupChain(s,chain.slice(0,chain.indexOf(scope)+1));if(member.groupInfo)placed.groupInfo=member.groupInfo.filter(g=>groupChain(placed).includes(g.id)).map(g=>({...g}));return placed;
}
