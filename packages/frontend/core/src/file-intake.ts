import {acquireIntake,releaseIntake} from './native-intake';
import {useEffect,useRef,useState} from 'react';
export function importableDrop(files:File[],insideEditor:boolean){return files.some(file=>/\.(md|markdown|zip)$/i.test(file.name)||!insideEditor&&/\.txt$/i.test(file.name));}
interface Options {ready:boolean;busy:boolean;onImport:(files:File[])=>void;notify:(message:string)=>void;}
const modalOpen=()=>!!document.querySelector('dialog[open],[role="dialog"][aria-modal="true"]');
export function useFileIntake(options:Options){
 const current=useRef(options);current.current=options;
 const [dragging,setDragging]=useState(false),wake=useRef<()=>void>(()=>{});
 useEffect(()=>{const bridge=window.mydocDesktop;if(!bridge)return;let alive=true,reading=false,wanted=true,version=0,awaitingDialog=false,dialogSeen=false;const token=Symbol();
  const pump=async()=>{if(!alive||reading||awaitingDialog||!wanted||!current.current.ready||current.current.busy||modalOpen()||!acquireIntake(token))return;reading=true;const revision=version;
   try{const batch=await bridge.takeOpenFiles();if(!alive)return;if(!batch){if(version===revision)wanted=false;return;}
    // Keep the batch queued if a modal appeared while the main process was reading.
    if(current.current.busy||modalOpen())return;
    if(batch.files.length){awaitingDialog=true;current.current.onImport(batch.files.map(file=>new File([file.bytes.slice().buffer],file.name,{type:/\.zip$/i.test(file.name)?'application/zip':'text/plain',lastModified:file.modified})));}
    if(batch.errors.length)current.current.notify(batch.errors.join('；'));
    await bridge.acknowledgeOpenFiles(batch.id);
   }catch{if(alive){wanted=false;current.current.notify('读取待打开文件失败，请通过“文件 → 打开文件”重试。');}}
   finally{if(!awaitingDialog)releaseIntake(token);reading=false;if(alive&&wanted&&!awaitingDialog)setTimeout(()=>void pump(),0);}
  };
  const request=()=>{wanted=true;version++;void pump();};wake.current=()=>void pump();
  const unsubscribe=bridge.onOpenFiles(request);
  const observer=new MutationObserver(()=>{if(awaitingDialog){if(modalOpen())dialogSeen=true;else if(dialogSeen){awaitingDialog=false;dialogSeen=false;releaseIntake(token);}}if(wanted)void pump();});observer.observe(document.body,{childList:true,subtree:true,attributes:true,attributeFilter:['open']});
  const resume=()=>void pump();window.addEventListener('mydoc:intake-ready',resume);
  void pump();return()=>{alive=false;unsubscribe();observer.disconnect();window.removeEventListener('mydoc:intake-ready',resume);wake.current=()=>{};releaseIntake(token);};
 },[]);
 useEffect(()=>{wake.current();},[options.ready,options.busy]);
 useEffect(()=>{let depth=0;const reset=()=>{depth=0;setDragging(false);};
  const enter=(event:DragEvent)=>{if(event.dataTransfer?.types.includes('Files')){depth++;if(!modalOpen())setDragging(true);}};
  const leave=()=>{if(--depth<=0)reset();};
  const over=(event:DragEvent)=>{if(event.dataTransfer?.types.includes('Files')){event.preventDefault();event.dataTransfer.dropEffect=current.current.busy?'none':'copy';}};
  const capture=(event:DragEvent)=>{if(!event.dataTransfer?.types.includes('Files'))return;reset();
   const target=event.target instanceof Element?event.target:null;
   if(modalOpen()){if(target?.closest('dialog[open],[role="dialog"][aria-modal="true"]')?.getAttribute('aria-label')==='导入文档')return;event.preventDefault();event.stopPropagation();if(!document.querySelector('dialog[open][aria-label="导入文档"]'))current.current.notify('请先完成或关闭当前弹窗，再拖入文件。');return;}
   if(!current.current.ready||current.current.busy){event.preventDefault();event.stopPropagation();current.current.notify('正在处理其他操作，请稍后重新拖入文件。');return;}
   if([...event.dataTransfer.items].some(item=>item.kind==='file'&&item.webkitGetAsEntry?.()?.isDirectory)){event.preventDefault();event.stopPropagation();current.current.notify('拖入文件夹请使用导入窗口的“选择文件夹”，或先压缩为 ZIP。');return;}
   const files=[...event.dataTransfer.files];if(importableDrop(files,!!target?.closest('.ProseMirror,.whiteboard-canvas'))){event.preventDefault();event.stopPropagation();current.current.onImport(files);}
  };
  const drop=(event:DragEvent)=>{if(!event.dataTransfer?.types.includes('Files')||event.defaultPrevented)return;event.preventDefault();const files=[...event.dataTransfer.files];if(!files.length)return;
   const host=document.querySelector('.workbench-pane.active-pane .whiteboard')??document.querySelector('.workbench-pane.active-pane .markdown-document');
   const request=new CustomEvent('mydoc:insert-files',{detail:files,cancelable:true});host?.dispatchEvent(request);
   if(!request.defaultPrevented)current.current.notify('请先打开文档，再把图片或附件拖入正文或白板。');
  };
  document.addEventListener('dragenter',enter);document.addEventListener('dragleave',leave);document.addEventListener('dragover',over);document.addEventListener('drop',capture,true);document.addEventListener('drop',drop);window.addEventListener('dragend',reset);window.addEventListener('blur',reset);
  return()=>{document.removeEventListener('dragenter',enter);document.removeEventListener('dragleave',leave);document.removeEventListener('dragover',over);document.removeEventListener('drop',capture,true);document.removeEventListener('drop',drop);window.removeEventListener('dragend',reset);window.removeEventListener('blur',reset);};
 },[]);
 return dragging;
}
