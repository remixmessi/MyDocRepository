import {useState} from 'react';
import type {WorkspaceAssets} from '@mydoc/nbstore';
import {retryImage} from './image-resources';
import {useImageResource} from './use-image-resource';
import {fullCrop} from './image-geometry';
import type {BoardImage} from './image-model';
import type {Viewport} from './model';
export function ImagePanel({image,view,assets,selected=false,anchor=false}:{image:BoardImage;view:Viewport;assets?:WorkspaceAssets;selected?:boolean;anchor?:boolean}) {
 const [attempt,setAttempt]=useState(0),{resource,error}=useImageResource(assets,image.assetId,attempt),crop=image.crop??fullCrop;
 return <div className="board-image" data-element-id={image.id} role="group" aria-label="白板图片" style={{width:image.width,height:image.height,transform:`translate(${view.x+image.x*view.zoom}px,${view.y+image.y*view.zoom}px) scale(${view.zoom})`}}><div className={`board-image-content ${selected?'selected':''} ${anchor?'anchored':''}`} style={{transform:`rotate(${image.rotation??0}deg)`}}><div className="board-image-clip">{resource&&!error?<img src={resource.url} alt={image.text||'图片'} draggable={false} style={{width:100/crop.width+'%',height:100/crop.height+'%',maxWidth:'none',transform:`translate(${-crop.x*100}%,${-crop.y*100}%)`}}/>:<div className="board-image-state" role="status"><span>{error||'正在读取图片…'}</span>{error&&<button onClick={()=>{if(assets)retryImage(assets,image.assetId);setAttempt(n=>n+1);}} aria-label={`重试图片 ${image.text}`}>重试</button>}</div>}</div>{selected&&<><span className="board-image-handle" style={{width:10/view.zoom,height:10/view.zoom,right:-5/view.zoom,bottom:-5/view.zoom,borderWidth:1/view.zoom}}/><span className="board-image-rotation" aria-hidden="true" style={{width:12/view.zoom,height:12/view.zoom,left:image.width/2-6/view.zoom,top:-36/view.zoom,borderWidth:1/view.zoom}}/></>}</div></div>;
}
