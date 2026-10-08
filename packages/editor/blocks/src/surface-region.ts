import {selectionBounds,type Bounds,type Shape,type SurfaceElement} from './whiteboard/model';
import {isHidden} from './whiteboard/status';
export function validRegion(value:unknown):value is Bounds{if(!value||typeof value!=='object')return false;const b=value as Bounds;return[b.x,b.y,b.width,b.height].every(v=>typeof v==='number'&&Number.isFinite(v)&&Math.abs(v)<=1e7)&&b.width>=1&&b.height>=1;}
export const regionKey=(region:Bounds)=>'@region:'+JSON.stringify([region.x,region.y,region.width,region.height]);
export function parseRegionKey(key:string):Bounds|undefined{if(!key.startsWith('@region:'))return;try{const [x,y,width,height]=JSON.parse(key.slice(8)),region={x,y,width,height};return validRegion(region)?region:undefined;}catch{return;}}
export function regionFrame(region:Bounds):Shape{return{id:regionKey(region),kind:'frame',x:region.x,y:region.y,width:region.width,height:region.height,text:'区域引用',fill:'#ffffff'};}
export const sceneRegion=(elements:SurfaceElement[])=>selectionBounds(elements.filter(s=>!isHidden(s)).map(s=>s.id),elements)??{x:0,y:0,width:600,height:400};
