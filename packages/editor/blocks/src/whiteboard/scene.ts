import {isHidden,isLocked} from './status';
import {elementBounds,hitElement,attachEndpoint,sortElements,type Bounds,type Point,type SurfaceElement,type Viewport} from './model';
import {groupChain,expandGroups} from './groups';
const CELL=256,MAX_CELLS=64,MAX_QUERY_CELLS=4096;
const intersects=(a:Bounds,b:Bounds)=>a.x<=b.x+b.width&&a.x+a.width>=b.x&&a.y<=b.y+b.height&&a.y+a.height>=b.y;
const range=(b:Bounds)=>[Math.floor(b.x/CELL),Math.floor(b.y/CELL),Math.floor((b.x+b.width)/CELL),Math.floor((b.y+b.height)/CELL)];
export class SurfaceScene {
  readonly ordered:SurfaceElement[];
  readonly byId:Map<string,SurfaceElement>;
  readonly bounds=new Map<string,Bounds>();
  private readonly ranks=new Map<string,number>();
  private readonly cells=new Map<string,Set<string>>();
  private readonly large=new Set<string>();
  readonly indexedElements:number;
  constructor(readonly elements:SurfaceElement[],previous?:SurfaceScene){
    if(previous){
      this.byId=new Map(elements.map(s=>[s.id,s]));
      const orderChanged=elements.length!==previous.elements.length||elements.some((s,i)=>s.id!==previous.elements[i]?.id||s.order!==previous.elements[i]?.order||s.kind!==previous.elements[i]?.kind);
      this.ordered=orderChanged?sortElements(elements):previous.ordered.map(s=>this.byId.get(s.id)!);this.ordered.forEach((s,i)=>this.ranks.set(s.id,i));
      this.cells=new Map(previous.cells);this.large=new Set(previous.large);this.bounds=new Map(previous.bounds);
      const changed=new Set<string>();for(const s of elements)if(previous.byId.get(s.id)!==s)changed.add(s.id);for(const id of previous.byId.keys())if(!this.byId.has(id))changed.add(id);
      for(const s of elements)if(s.kind==='connector'&&([s.source,s.target].some(p=>p.anchor&&changed.has(p.anchor.id))||changed.size>0&&s.routing==='orthogonal'&&s.avoid!==false))changed.add(s.id);
      const touched=new Set<string>();const bucket=(key:string)=>{if(!touched.has(key)){this.cells.set(key,new Set(this.cells.get(key)));touched.add(key);}return this.cells.get(key)!;};
      for(const id of changed){const old=previous.bounds.get(id);if(old&&!previous.large.has(id)){const [x,y,right,bottom]=range(old);for(let a=x;a<=right;a++)for(let b=y;b<=bottom;b++)bucket(a+':'+b).delete(id);}this.large.delete(id);this.bounds.delete(id);const element=this.byId.get(id);if(!element)continue;const bound=elementBounds(element,elements,this.byId);this.bounds.set(id,bound);const [x,y,right,bottom]=range(bound);if((right-x+1)*(bottom-y+1)>MAX_CELLS)this.large.add(id);else for(let a=x;a<=right;a++)for(let b=y;b<=bottom;b++)bucket(a+':'+b).add(id);}
      for(const key of touched)if(!this.cells.get(key)!.size)this.cells.delete(key);this.indexedElements=changed.size;return;
    }
    this.indexedElements=elements.length;
    this.ordered=sortElements(elements);this.byId=new Map(elements.map(s=>[s.id,s]));
    this.ordered.forEach((s,i)=>{this.ranks.set(s.id,i);const b=elementBounds(s,elements,this.byId);this.bounds.set(s.id,b);const [x,y,right,bottom]=range(b);
      if((right-x+1)*(bottom-y+1)>MAX_CELLS){this.large.add(s.id);return;}
      for(let a=x;a<=right;a++)for(let b=y;b<=bottom;b++){const key=a+':'+b;if(!this.cells.has(key))this.cells.set(key,new Set());this.cells.get(key)!.add(s.id);}
    });
  }
  orderedIds(ids:string[]){return [...new Set(ids)].filter(id=>this.byId.has(id)).sort((a,b)=>this.ranks.get(a)!-this.ranks.get(b)!).map(id=>this.byId.get(id)!);}
  query(box:Bounds){
    const [x,y,right,bottom]=range(box);let ids:Set<string>;
    // Very large elements/queries use a bounded fallback instead of allocating millions of cells.
    if((right-x+1)*(bottom-y+1)>MAX_QUERY_CELLS)ids=new Set(this.byId.keys());
    else{ids=new Set(this.large);for(let a=x;a<=right;a++)for(let b=y;b<=bottom;b++)for(const id of this.cells.get(a+':'+b)??[])ids.add(id);}
    return [...ids].filter(id=>intersects(this.bounds.get(id)!,box)).sort((a,b)=>this.ranks.get(a)!-this.ranks.get(b)!).map(id=>this.byId.get(id)!);
  }
  viewport(view:Viewport,width:number,height:number,padding=32){const pad=Math.max(32,padding/view.zoom);return this.query({x:-view.x/view.zoom-pad,y:-view.y/view.zoom-pad,width:width/view.zoom+2*pad,height:height/view.zoom+2*pad});}
  hit(point:Point,tolerance:number,scope?:string){
    const pad=Math.max(28,tolerance+6),candidates=this.query({x:point.x-pad,y:point.y-pad,width:pad*2,height:pad*2});
    for(let i=candidates.length-1;i>=0;i--){const s=candidates[i];if(!isHidden(s)&&!isLocked(s)&&(!scope||groupChain(s).includes(scope))&&hitElement(s,point,this.elements,tolerance,this.byId))return s;}
  }
  attach(point:Point,tolerance:number){return attachEndpoint(point,this.query({x:point.x-tolerance,y:point.y-tolerance,width:2*tolerance,height:2*tolerance}),tolerance);}
  box(box:Bounds,scope?:string){return expandGroups(this.query(box).filter(s=>!isHidden(s)&&!isLocked(s)).filter(s=>{const b=this.bounds.get(s.id)!;return b.x>=box.x&&b.y>=box.y&&b.x+b.width<=box.x+box.width&&b.y+b.height<=box.y+box.height;}).map(s=>s.id),this.elements,scope);}
}
const cache=new WeakMap<SurfaceElement[],SurfaceScene>();
// Scene arrays and their elements are immutable snapshots, including gesture previews.
export function sceneFor(elements:SurfaceElement[],previous?:SurfaceScene){let scene=cache.get(elements);if(!scene){scene=new SurfaceScene(elements,previous);cache.set(elements,scene);}return scene;}
