import {validCrop,type ImageCrop} from './image-geometry';
import type {AssetRecord} from '@mydoc/nbstore';
import type {Point,Shape} from './model';
export interface BoardImage extends Omit<Shape,'kind'> {kind:'image';assetId:string;pixelWidth:number;pixelHeight:number;rotation?:number;crop?:ImageCrop}
export function validImage(value:unknown):value is BoardImage {
  const image=value as BoardImage;
  return (image.rotation===undefined||Number.isFinite(image.rotation)&&image.rotation>=0&&image.rotation<360)&&(image.crop===undefined||validCrop(image.crop))&&image?.kind==='image'&&typeof image.assetId==='string'&&!!image.assetId&&image.assetId.length<=200&&[image.pixelWidth,image.pixelHeight].every(n=>Number.isInteger(n)&&n>=1&&n<=1e7);
}
export function createImage(asset:Pick<AssetRecord,'id'|'name'>,pixels:{width:number;height:number},center:Point):BoardImage {
  const scale=Math.max(1/pixels.width,1/pixels.height,Math.min(480/pixels.width,360/pixels.height)),width=Math.max(1,pixels.width*scale),height=Math.max(1,pixels.height*scale);
  return {id:crypto.randomUUID(),kind:'image',assetId:asset.id,pixelWidth:pixels.width,pixelHeight:pixels.height,x:center.x-width/2,y:center.y-height/2,width,height,text:asset.name.slice(0,2000)||'图片',fill:'#ffffff'};
}
export function resizeImage(image:BoardImage,width:number,height:number,axis:'width'|'height'='width'):BoardImage {
  const factor=Math.max(1/image.width,1/image.height,axis==='width'?width/image.width:height/image.height);
  return {...image,width:image.width*factor,height:image.height*factor};
}
