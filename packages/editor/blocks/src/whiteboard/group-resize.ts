import {isLocked} from './status';
import{expandGroups,expandFrames,selectionBounds,resolveEndpoint,validSurface,type SurfaceElement}from'./model';
import{rebuildMindmap}from'./mindmap';
export function resizeSelection(ids:string[],elements:SurfaceElement[],width:number,height:number,keepRatio=false,scope?:string){
  if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)throw new Error('组合尺寸必须是正数');
  const members=expandFrames(expandGroups(ids,elements,scope),elements),selected=new Set(members),bounds=selectionBounds(members,elements);
  if(!bounds||members.length<2)throw new Error('请先选择至少两个元素');
  const items=elements.filter(s=>selected.has(s.id));if(items.some(isLocked))throw new Error('所选内容已锁定');const byId=new Map(elements.map(s=>[s.id,s]));let sx=width/bounds.width,sy=height/bounds.height,minX=.001,minY=.001,max=Infinity;
  const uniform=keepRatio||items.some(s=>s.kind==='mindmap'||s.kind==='image');
  for(const s of items){if(s.kind==='connector'||s.kind==='brush')continue;if(s.kind==='mindmap'){minX=Math.max(minX,.1/(s.scale??1));minY=Math.max(minY,.1/(s.scale??1));max=Math.min(max,4/(s.scale??1));continue;}const w=s.kind==='note'?240:s.kind==='doc-card'?160:s.kind==='image'?1:24,h=s.kind==='note'?160:s.kind==='doc-card'?(s.presentation==='card'?100:56):s.kind==='image'?1:24;minX=Math.max(minX,w/s.width);minY=Math.max(minY,h/s.height);}
  if(uniform){const desired=Math.abs(sx-1)>=Math.abs(sy-1)?sx:sy;const factor=Math.max(minX,minY,Math.min(max,desired));sx=factor;sy=factor;}else{sx=Math.max(minX,sx);sy=Math.max(minY,sy);}
  const transform=(p:{x:number;y:number})=>({x:bounds.x+(p.x-bounds.x)*sx,y:bounds.y+(p.y-bounds.y)*sy});
  let changes:SurfaceElement[]=items.map(s=>{const p=transform(s);if(s.kind==='connector'){const endpoint=(key:'source'|'target')=>s[key].anchor&&!selected.has(s[key].anchor!.id)?s[key]:{...s[key],...transform(resolveEndpoint(s[key],elements,byId))};return{...s,...p,source:endpoint('source'),target:endpoint('target'),...(s.waypoints?{waypoints:s.waypoints.map(transform)}:{}),...(s.bend!==undefined?{bend:s.bend*Math.max(sx,sy)}:{})};}if(s.kind==='mindmap')return rebuildMindmap({...s,...p,scale:Math.max(.1,Math.min(4,(s.scale??1)*sx))});const minW=s.kind==='note'?240:s.kind==='doc-card'?160:s.kind==='brush'||s.kind==='image'?1:24,minH=s.kind==='note'?160:s.kind==='doc-card'?(s.presentation==='card'?100:56):s.kind==='brush'||s.kind==='image'?1:24;return{...s,...p,width:Math.max(minW,s.width*sx),height:Math.max(minH,s.height*sy),...(s.kind==='brush'?{points:s.points.map(p=>({...p,x:p.x*sx,y:p.y*sy}))}:{})};});
  const changed=new Map(changes.map(s=>[s.id,s])),scene=elements.map(s=>changed.get(s.id)??s),sceneById=new Map(scene.map(s=>[s.id,s]));
  changes=changes.map(s=>{if(s.kind!=='connector')return s;const a=resolveEndpoint(s.source,scene,sceneById),b=resolveEndpoint(s.target,scene,sceneById);return{...s,x:Math.min(a.x,b.x),y:Math.min(a.y,b.y),width:Math.max(1,Math.abs(a.x-b.x)),height:Math.max(1,Math.abs(a.y-b.y))};});
  changes.forEach(s=>sceneById.set(s.id,s));if(!validSurface(sceneById))throw new Error('缩放超出元素范围，已保留原内容');
  return {changes,width:bounds.width*sx,height:bounds.height*sy,uniform};
}
