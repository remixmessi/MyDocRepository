import type {BoardImage} from './image-model';
import type {Point,Bounds} from './model';
export interface ImageCrop {x:number;y:number;width:number;height:number}
export const fullCrop:ImageCrop={x:0,y:0,width:1,height:1};
export function validCrop(c:ImageCrop){return !!c&&[c.x,c.y,c.width,c.height].every(Number.isFinite)&&c.x>=0&&c.y>=0&&c.width>0&&c.height>0&&c.x+c.width<=1+1e-9&&c.y+c.height<=1+1e-9;}
export function rotateVector(p:Point,degrees:number):Point {const a=degrees*Math.PI/180;return {x:p.x*Math.cos(a)-p.y*Math.sin(a),y:p.x*Math.sin(a)+p.y*Math.cos(a)};}
export function imagePoint(image:BoardImage,p:Point,inverse=false):Point {const cx=image.x+image.width/2,cy=image.y+image.height/2,v=rotateVector({x:p.x-cx,y:p.y-cy},(inverse?-1:1)*(image.rotation??0));return {x:cx+v.x,y:cy+v.y};}
export function imageBounds(image:BoardImage):Bounds {const corners=[{x:image.x,y:image.y},{x:image.x+image.width,y:image.y},{x:image.x,y:image.y+image.height},{x:image.x+image.width,y:image.y+image.height}].map(p=>imagePoint(image,p));const x=Math.min(...corners.map(p=>p.x)),y=Math.min(...corners.map(p=>p.y));return{x,y,width:Math.max(...corners.map(p=>p.x))-x,height:Math.max(...corners.map(p=>p.y))-y};}
export const rotationHandle=(image:BoardImage,zoom:number)=>imagePoint(image,{x:image.x+image.width/2,y:image.y-30/zoom});
export function cropImage(image:BoardImage,crop:ImageCrop):BoardImage {if(!validCrop(crop))throw new Error('裁剪范围必须位于原图内');const old=image.crop??fullCrop,width=image.width/old.width*crop.width,height=image.height/old.height*crop.height;if(width<1||height<1)throw new Error('裁剪后尺寸不能小于 1');return {...image,crop:{...crop},x:image.x+(image.width-width)/2,y:image.y+(image.height-height)/2,width,height};}
export function resizeRotatedImage(image:BoardImage,next:BoardImage):BoardImage {const tl=imagePoint(image,image),v=rotateVector({x:next.width/2,y:next.height/2},image.rotation??0);return {...next,x:tl.x+v.x-next.width/2,y:tl.y+v.y-next.height/2};}
