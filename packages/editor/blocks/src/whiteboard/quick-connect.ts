import {isLocked,isHidden} from './status';
import {isAnchorShape,type SurfaceElement,type Shape,type Connector} from './model';
export function quickConnect(source:SurfaceElement,elements:SurfaceElement[],direction:'right'|'left'|'up'|'down'='right'){
 if(!isAnchorShape(source)||isLocked(source)||isHidden(source))throw new Error('请选择未锁定的可连接元素');
 const horizontal=direction==='right'||direction==='left',positive=direction==='right'||direction==='down',width=160,height=96,center={x:source.x+source.width/2,y:source.y+source.height/2};
 const order=Math.max(-1,...elements.map((e,i)=>e.order??i))+1;
 const shape:Shape={id:crypto.randomUUID(),kind:'rectangle',text:'新图形',fill:source.fill,width,height,x:horizontal?(positive?source.x+source.width+96:source.x-width-96):center.x-width/2,y:horizontal?center.y-height/2:positive?source.y+source.height+96:source.y-height-96,groupId:source.groupId,groupPath:source.groupPath,groupInfo:source.groupInfo,order};
 const u=horizontal?(positive?1:0):.5,v=horizontal?.5:positive?1:0;
 const line:Connector={id:crypto.randomUUID(),kind:'connector',x:center.x,y:center.y,width:1,height:1,text:'',fill:'#ffffff',color:'#20242c',lineWidth:2,arrow:true,routing:'orthogonal',source:{x:center.x,y:center.y,anchor:{id:source.id,u,v}},target:{x:shape.x+width/2,y:shape.y+height/2,anchor:{id:shape.id,u:horizontal?1-u:.5,v:horizontal?.5:1-v}},groupId:source.groupId,groupPath:source.groupPath,groupInfo:source.groupInfo,order:order+1};
 return {shape,line};
}
