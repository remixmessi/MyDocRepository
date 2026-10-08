import {useEffect,useState} from 'react';
import type {WorkspaceAssets} from '@mydoc/nbstore';
import {acquireImage,type ImageResource} from './image-resources';
export function useSceneImages(assets:WorkspaceAssets|undefined,ids:string[]){const [images,setImages]=useState<ReadonlyMap<string,ImageResource>>(new Map());const key=JSON.stringify(ids);useEffect(()=>{let disposed=false;const ready=new Map<string,ImageResource>();setImages(ready);if(!assets)return;const leases=ids.map(id=>{const lease=acquireImage(assets,id);void lease.promise.then(resource=>{if(!disposed){ready.set(id,resource);setImages(new Map(ready));}},()=>{});return lease;});return()=>{disposed=true;leases.forEach(lease=>lease.release());};},[assets,key]);return images;}
