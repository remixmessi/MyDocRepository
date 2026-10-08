import {useEffect,useRef,useState} from 'react';
import {MAX_ASSET_SIZE,type WorkspaceAssets} from '@mydoc/nbstore';
import {createImage,type BoardImage} from './image-model';
import {imagePixels} from './image-import';
import {placeInScope,type SurfaceStore,type Point} from './model';
export function useImageImports(surface:SurfaceStore,assets:WorkspaceAssets|undefined,scope:string|undefined,selected:(ids:string[])=>void,notify:(text:string)=>void) {
  const [progress,setProgress]=useState(''),busy=useRef(false),alive=useRef(true);
  useEffect(()=>{alive.current=true;return()=>{alive.current=false;};},[surface]);
  async function add(files:File[],center:Point,existingId?:string) {
    if(busy.current||!files.length&&!existingId)return;if(!assets){notify('图片存储不可用');return;}if(files.length>10){notify('每次最多添加 10 张图片');return;}if(surface.elements.size+(existingId?1:files.length)>5000){notify('白板最多支持 5,000 个元素');return;}
    busy.current=true;notify('');const images:BoardImage[]=[];const failures:string[]=[];
    try {
      const inputs: (File|string)[]=existingId?[existingId]:files;
      for(const [index,input] of inputs.entries()) {
        if(!alive.current)return;setProgress(`正在添加图片 ${index+1}/${inputs.length}…`);
        try {
          let asset,pixels;
          if(typeof input==='string'){asset=await assets.get(input);if(!asset)throw new Error('本地图片缺失，请恢复包含文件的工作区备份');pixels=await imagePixels(asset.bytes);}
          else{if(input.size>MAX_ASSET_SIZE)throw new Error('单个文件不能超过 20 MB');pixels=await imagePixels(new Uint8Array(await input.arrayBuffer()));if(!alive.current)return;asset=await assets.save(input,true);}
          if(!alive.current)return;images.push(createImage(asset,pixels,{x:center.x+index*24,y:center.y+index*24}));
        }catch(err){failures.push(`${typeof input==='string'?'正文图片':input.name}：${err instanceof Error?err.message:'图片保存失败，请重试'}`);}
      }
      if(alive.current&&images.length){const all=surface.list(),order=Math.max(-1,...all.map((s,i)=>s.order??i))+1;surface.apply(images.map((s,i)=>placeInScope({...s,order:order+i},all,scope)));selected(images.map(s=>s.id));}
    }catch(err){failures.push(err instanceof Error?err.message:'图片保存失败');}
    finally{busy.current=false;if(alive.current){setProgress('');notify(failures.join('；'));}}
  }
  return {progress,add};
}
