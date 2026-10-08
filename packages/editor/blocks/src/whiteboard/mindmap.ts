import type { Shape, Point } from './model';
export type MindDirection='left'|'right'|'both';
export interface MindNode { id:string; parent?:string; text:string; collapsed?:boolean }
export interface Mindmap extends Omit<Shape,'kind'> {kind:'mindmap'; nodes:MindNode[]; direction:MindDirection; scale?:number}
export interface NodeBox extends MindNode, Point {width:number;height:number;descendants:number}
export function validMindmap(value:unknown):value is Mindmap {
  const m=value as Mindmap;if(!m||m.kind!=='mindmap'||!['left','right','both'].includes(m.direction)||!Array.isArray(m.nodes)||!m.nodes.length||m.nodes.length>200)return false;
  if(m.scale!==undefined&&(!Number.isFinite(m.scale)||m.scale<.1||m.scale>4))return false;
  const ids=new Set<string>();for(const n of m.nodes){if(!n||typeof n.id!=='string'||!n.id||n.id.length>200||ids.has(n.id)||typeof n.text!=='string'||n.text.length>200||n.parent!==undefined&&(typeof n.parent!=='string'||!n.parent))return false;if(n.collapsed!==undefined&&typeof n.collapsed!=='boolean')return false;ids.add(n.id);}
  const roots=m.nodes.filter(n=>n.parent===undefined);if(roots.length!==1)return false;
  const byId=new Map(m.nodes.map(n=>[n.id,n]));for(const n of m.nodes){const seen=new Set<string>();let current:MindNode|undefined=n;while(current){if(seen.has(current.id)||seen.size>=50)return false;seen.add(current.id);if(current.parent&&!byId.has(current.parent))return false;current=current.parent?byId.get(current.parent):undefined;}}
  const layout=layoutMindmap(m.nodes,m.direction);return m.width===layout.width*(m.scale??1)&&m.height===layout.height*(m.scale??1)&&m.text===roots[0].text;
}
// Collapsed nodes retain their descendants in the model but allocate no space for them.
export function layoutMindmap(nodes:MindNode[],direction:MindDirection) {
  const children=(id:string)=>nodes.filter(n=>n.parent===id),spans=new Map<string,number>(),counts=new Map<string,number>(),positions:NodeBox[]=[];
  const root=nodes.find(n=>n.parent===undefined)!;
  const count=(n:MindNode):number=>{const value=children(n.id).reduce((sum,k)=>sum+1+count(k),0);counts.set(n.id,value);return value;};count(root);
  const visibleChildren=(n:MindNode)=>n.collapsed?[]:children(n.id);
  const measure=(n:MindNode):number=>{const kids=visibleChildren(n);const span=Math.max(64,kids.reduce((sum,k)=>sum+measure(k),0)+Math.max(0,kids.length-1)*24);spans.set(n.id,span);return span;};
  let height=measure(root);
  const place=(n:MindNode,depth:number,top:number,sign:number)=>{positions.push({...n,x:sign*depth*224,y:top+(spans.get(n.id)!-64)/2,width:160,height:64,descendants:counts.get(n.id)!});let cursor=top;for(const child of visibleChildren(n)){place(child,depth+1,cursor,sign);cursor+=spans.get(child.id)!+24;}};
  if(direction==='both'){
    const kids=visibleChildren(root),right=kids.filter((_,i)=>i%2===0),left=kids.filter((_,i)=>i%2===1);
    const span=(group:MindNode[])=>group.reduce((sum,n)=>sum+spans.get(n.id)!,0)+Math.max(0,group.length-1)*24;
    height=Math.max(64,span(left),span(right));
    positions.push({...root,x:0,y:(height-64)/2,width:160,height:64,descendants:counts.get(root.id)!});
    for(const [group,sign] of [[right,1],[left,-1]] as const){let cursor=(height-span(group))/2;for(const n of group){place(n,1,cursor,sign);cursor+=spans.get(n.id)!+24;}}
  }else place(root,0,0,direction==='left'?-1:1);
  const minX=Math.min(...positions.map(n=>n.x)),maxX=Math.max(...positions.map(n=>n.x+160));
  return {width:maxX-minX,height,boxes:positions.map(n=>({...n,x:n.x-minX}))};
}
export function toggleMindNode(m:Mindmap,id:string){const node=m.nodes.find(n=>n.id===id);if(!node)throw new Error('请选择主题');if(!m.nodes.some(n=>n.parent===id))return m;return rebuildMindmap(m,m.nodes.map(n=>n.id===id?{...n,collapsed:!n.collapsed}:n));}
export function visibleMindNode(m:Mindmap,id?:string){const boxes=layoutMindmap(m.nodes,m.direction).boxes,visible=new Set(boxes.map(n=>n.id));let node=m.nodes.find(n=>n.id===id);while(node&&!visible.has(node.id))node=m.nodes.find(n=>n.id===node!.parent);return node?.id??boxes.find(n=>!n.parent)!.id;}
export function hitMindToggle(m:Mindmap,p:Point){return layoutMindmap(m.nodes,m.direction).boxes.find(n=>n.descendants>0&&(p.x-m.x)/(m.scale??1)>=n.x+126&&(p.x-m.x)/(m.scale??1)<=n.x+156&&(p.y-m.y)/(m.scale??1)>=n.y+38&&(p.y-m.y)/(m.scale??1)<=n.y+62)?.id;}
export function rebuildMindmap(m:Mindmap,nodes=m.nodes,direction=m.direction):Mindmap {const layout=layoutMindmap(nodes,direction);return{...m,nodes,direction,text:nodes.find(n=>n.parent===undefined)!.text,width:layout.width*(m.scale??1),height:layout.height*(m.scale??1)};}
export function createMindmap(p:Point):Mindmap {return rebuildMindmap({id:crypto.randomUUID(),kind:'mindmap',...p,width:160,height:64,text:'中心主题',fill:'#e8efff',direction:'right',nodes:[{id:crypto.randomUUID(),text:'中心主题'}]});}
export function addMindNode(m:Mindmap,id:string,sibling=false) {const current=m.nodes.find(n=>n.id===id);if(!current)throw new Error('请选择主题');if(sibling&&!current.parent)throw new Error('中心主题没有同级主题');if(m.nodes.length>=200)throw new Error('每张思维导图最多支持 200 个主题');const node:MindNode={id:crypto.randomUUID(),parent:sibling?current.parent:current.id,text:'新主题'};const nodes=[...m.nodes];let index=nodes.length;if(sibling){index=nodes.findIndex(n=>n.id===id)+1;}nodes.splice(index,0,node);const ancestors=new Set<string>();let parent=node.parent;while(parent){ancestors.add(parent);parent=nodes.find(n=>n.id===parent)?.parent;}for(let i=0;i<nodes.length;i++)if(ancestors.has(nodes[i].id)&&nodes[i].collapsed)nodes[i]={...nodes[i],collapsed:false};const result=rebuildMindmap(m,nodes);if(!validMindmap(result))throw new Error('主题层级最多支持 50 层');return{map:result,id:node.id};}
export function removeMindNode(m:Mindmap,id:string){const current=m.nodes.find(n=>n.id===id);if(!current?.parent)throw new Error('中心主题不能删除，请删除整张导图');const removed=new Set([id]);for(let i=0;i<m.nodes.length;i++)for(const n of m.nodes)if(n.parent&&removed.has(n.parent))removed.add(n.id);return rebuildMindmap(m,m.nodes.filter(n=>!removed.has(n.id)));}
export function hitMindNode(m:Mindmap,p:Point){return layoutMindmap(m.nodes,m.direction).boxes.find(n=>(p.x-m.x)/(m.scale??1)>=n.x&&(p.x-m.x)/(m.scale??1)<=n.x+n.width&&(p.y-m.y)/(m.scale??1)>=n.y&&(p.y-m.y)/(m.scale??1)<=n.y+n.height)?.id;}
export function drawMindmap(ctx:CanvasRenderingContext2D,m:Mindmap,zoom:number,selected:boolean,nodeId?:string){
  const {boxes}=layoutMindmap(m.nodes,m.direction);nodeId=boxes.some(n=>n.id===nodeId)?nodeId:boxes.find(n=>!n.parent)!.id;ctx.save();ctx.translate(m.x,m.y);const scale=m.scale??1;ctx.scale(scale,scale);zoom*=scale;ctx.strokeStyle='#8c9fc6';ctx.lineWidth=2;
  for(const n of boxes){const parent=boxes.find(p=>p.id===n.parent);if(!parent)continue;const right=n.x>parent.x,a={x:parent.x+(right?160:0),y:parent.y+32},b={x:n.x+(right?0:160),y:n.y+32},mid=(a.x+b.x)/2;ctx.beginPath();ctx.moveTo(a.x,a.y);ctx.bezierCurveTo(mid,a.y,mid,b.y,b.x,b.y);ctx.stroke();}
  for(const n of boxes){ctx.fillStyle=n.parent?m.fill:'#174bd6';ctx.strokeStyle=selected&&n.id===nodeId?'#174bd6':'#9baac5';ctx.lineWidth=(selected&&n.id===nodeId?3:1)/zoom;ctx.beginPath();ctx.roundRect(n.x,n.y,160,64,8);ctx.fill();ctx.stroke();ctx.save();ctx.beginPath();ctx.rect(n.x+12,n.y+8,n.descendants?110:136,48);ctx.clip();ctx.fillStyle=n.parent?'#20242c':'#ffffff';ctx.font='15px system-ui';ctx.textBaseline='top';let line='',lines:string[]=[];for(const c of n.text||'未命名主题'){if(c==='\n'||ctx.measureText(line+c).width>(n.descendants?110:136)){lines.push(line);line='';if(c==='\n')continue;}line+=c;}lines.push(line);lines.slice(0,2).forEach((line,i)=>ctx.fillText(line,n.x+12,n.y+12+i*22));ctx.restore();if(n.descendants){ctx.fillStyle='#ffffff';ctx.strokeStyle='#9baac5';ctx.lineWidth=1/zoom;ctx.beginPath();ctx.roundRect(n.x+126,n.y+40,30,20,4);ctx.fill();ctx.stroke();ctx.fillStyle='#174bd6';ctx.font='10px system-ui';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(n.collapsed?'+'+n.descendants:'−',n.x+141,n.y+50);ctx.textAlign='left';ctx.textBaseline='top';}}
  if(selected){ctx.strokeStyle='#174bd6';ctx.lineWidth=1/zoom;ctx.setLineDash([5/zoom,4/zoom]);ctx.strokeRect(-4/zoom,-4/zoom,m.width/scale+8/zoom,m.height/scale+8/zoom);}ctx.restore();
}

export function mindDescendants(m:Mindmap,id:string){const ids=new Set([id]);for(let i=0;i<m.nodes.length;i++)for(const n of m.nodes)if(n.parent&&ids.has(n.parent))ids.add(n.id);return ids;}
export function reparentMindNode(m:Mindmap,id:string,parentId:string):Mindmap{
  const node=m.nodes.find(n=>n.id===id),parent=m.nodes.find(n=>n.id===parentId);
  if(!node||!parent)throw new Error('主题已不存在，请重新选择');
  if(!node.parent)throw new Error('中心主题不能更换父级');
  const branch=mindDescendants(m,id);
  if(branch.has(parentId))throw new Error('不能移入自身或后代主题');
  if(node.parent===parentId)return m;
  const ancestors=new Set<string>();let current:MindNode|undefined=parent;while(current){ancestors.add(current.id);current=m.nodes.find(n=>n.id===current!.parent);}
  const nodes=[...m.nodes.filter(n=>!branch.has(n.id)),...m.nodes.filter(n=>branch.has(n.id))].map(n=>n.id===id?{...n,parent:parentId}:ancestors.has(n.id)&&n.collapsed?{...n,collapsed:false}:n);
  const next=rebuildMindmap(m,nodes);
  if(!validMindmap(next))throw new Error('换父级后主题层级超过 50 层');
  return next;
}
