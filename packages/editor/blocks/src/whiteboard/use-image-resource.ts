import {useEffect,useState} from 'react';
import type {WorkspaceAssets} from '@mydoc/nbstore';
import {acquireImage,type ImageResource} from './image-resources';
export function useImageResource(assets:WorkspaceAssets|undefined,id:string,attempt=0){const [state,setState]=useState<{resource?:ImageResource;error?:string}>({});useEffect(()=>{let disposed=false;setState({});if(!assets){setState({error:'图片存储不可用'});return;}const lease=acquireImage(assets,id);void lease.promise.then(resource=>{if(!disposed)setState({resource});},err=>{if(!disposed)setState({error:err instanceof Error?err.message:'图片读取失败'});});return()=>{disposed=true;lease.release();};},[assets,id,attempt]);return state;}
