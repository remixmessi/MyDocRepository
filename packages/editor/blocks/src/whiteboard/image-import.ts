import {imageMime,type AssetRecord} from '@mydoc/nbstore';
export async function imagePixels(bytes:Uint8Array):Promise<{width:number;height:number}> {
  const mime=imageMime(bytes);if(!mime)throw new Error('图片支持 PNG、JPEG、GIF、WebP');
  const url=URL.createObjectURL(new Blob([bytes],{type:mime}));
  try{return await new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>img.naturalWidth&&img.naturalHeight?resolve({width:img.naturalWidth,height:img.naturalHeight}):reject(new Error('图片尺寸无法读取'));img.onerror=()=>reject(new Error('图片数据损坏，无法读取'));img.src=url;});}finally{URL.revokeObjectURL(url);}
}
export function downloadImage(asset:AssetRecord) {const url=URL.createObjectURL(new Blob([asset.bytes],{type:asset.mime})),link=document.createElement('a');link.href=url;link.download=asset.name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
