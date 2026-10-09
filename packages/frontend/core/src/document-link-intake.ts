import {useEffect,useRef} from 'react';
import type {DocumentLocation} from '../../../common/nbstore/src/deep-link';
import {acquireIntake,releaseIntake} from './native-intake';
export function useDocumentLinkIntake(options:{ready:boolean;busy:boolean;open:(target:DocumentLocation)=>Promise<void>;notify:(message:string)=>void}){
 const current=useRef(options);current.current=options;const wake=useRef(()=>{});
 useEffect(()=>{const bridge=window.mydocDesktop;if(!bridge)return;let alive=true,reading=false,wanted=true,version=0;const token=Symbol();
  const modal=()=>!!document.querySelector('dialog[open],[role="dialog"][aria-modal="true"]');
  const pump=async()=>{if(!alive||!wanted||reading||!current.current.ready||current.current.busy||modal()||!acquireIntake(token))return;reading=true;const revision=version;let consumed=false;
   try{const request=await bridge.takeDocumentLink();if(!alive)return;if(!request){if(revision===version)wanted=false;return;}if(current.current.busy||modal())return;
    try{if(request.error)throw new Error(request.error);if(request.target)await current.current.open(request.target);}catch(error){current.current.notify(error instanceof Error?error.message:'无法打开文档链接，请重试。');}
    await bridge.acknowledgeDocumentLink(request.id);consumed=true;
   }catch{wanted=false;if(alive)current.current.notify('读取桌面文档链接失败，请重新打开链接。');}
   finally{releaseIntake(token);reading=false;if(alive&&consumed)setTimeout(()=>void pump(),0);}
  };
  wake.current=()=>void pump();const unsub=bridge.onDocumentLink(()=>{wanted=true;version++;void pump();});
  // Acknowledgement wakes the next request; modal closure and file intake release resume deferred links.
  const resume=()=>{if(!reading)void pump();};window.addEventListener('mydoc:intake-ready',resume);
  const observer=new MutationObserver(resume);observer.observe(document.body,{subtree:true,childList:true,attributes:true,attributeFilter:['open']});void pump();
  return()=>{alive=false;unsub();observer.disconnect();window.removeEventListener('mydoc:intake-ready',resume);wake.current=()=>{};releaseIntake(token);};
 },[]);
 useEffect(()=>wake.current(),[options.ready,options.busy]);
}
