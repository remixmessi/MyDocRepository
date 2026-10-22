import {acquireIntake,releaseIntake} from './native-intake';
import {useEffect,useRef,useState} from 'react';
import {officeImportable} from '../../../common/nbstore/src/office';
export function importableDrop(files:File[],insideEditor:boolean){return files.some(file=>/\.(md|markdown|zip)$/i.test(file.name)||!insideEditor&&/\.txt$/i.test(file.name));}
/** Office files dropped outside the editor body import as files; PDFs alone stay attachments while a document is open. */
const LEGACY_OFFICE=/\.(doc|xls|ppt|dot|xlt|pot)$/i;
export function officeDrop(files:File[],insideEditor:boolean){return !insideEditor&&files.some(file=>officeImportable(file.name)&&!/\.pdf$/i.test(file.name)||LEGACY_OFFICE.test(file.name));}
export type JsonImportKind='workspace-backup'|'draft'|'document-template'|'office-backup'|'whiteboard-template';
export function jsonImportKind(text:string):JsonImportKind|undefined{
 let value:unknown;try{value=JSON.parse(text);}catch{return undefined;}
 if(!value||typeof value!=='object'||Array.isArray(value))return undefined;const data=value as {type?:unknown;format?:unknown};
 if(data.type==='mydoc-workspace'||data.type==='mydoc-team-workspace')return 'workspace-backup';
 if(data.format==='mydoc-workspace-draft'||data.format==='mydoc-document-draft')return 'draft';
 if(data.type==='mydoc-document-template')return 'document-template';
 if(data.type==='mydoc-office-resource')return 'office-backup';
 if(data.type==='mydoc-whiteboard-template')return 'whiteboard-template';
 return undefined;
}
interface Options {ready:boolean;busy:boolean;onImport:(files:File[])=>void;onOfficeImport:(files:File[])=>boolean;onJsonImport:(file:File,kind:JsonImportKind)=>void;notify:(message:string)=>void;}
const modalOpen=()=>!!document.querySelector('dialog[open],[role="dialog"][aria-modal="true"]');
const MODAL='dialog[open],[role="dialog"][aria-modal="true"]';
export function useFileIntake(options:Options){
 const current=useRef(options);current.current=options;
 const [dragging,setDragging]=useState(false),wake=useRef<()=>void>(()=>{});
 useEffect(()=>{const bridge=window.mydocDesktop;if(!bridge)return;let alive=true,reading=false,wanted=true,version=0,awaitingDialog=false,dialogSeen=false;const token=Symbol();
  const pump=async()=>{if(!alive||reading||awaitingDialog||!wanted||!current.current.ready||current.current.busy||modalOpen()||!acquireIntake(token))return;reading=true;const revision=version;
   try{const batch=await bridge.takeOpenFiles();if(!alive)return;if(!batch){if(version===revision)wanted=false;return;}
    // Keep the batch queued if a modal appeared while the main process was reading.
    if(current.current.busy||modalOpen())return;
    if(batch.files.length){awaitingDialog=true;const files=batch.files.map(file=>new File([file.bytes.slice().buffer],file.name,{type:/\.zip$/i.test(file.name)?'application/zip':officeImportable(file.name)?'application/octet-stream':'text/plain',lastModified:file.modified}));
     // The main process queues Office files and Markdown/ZIP in separate batches.
     if(!files.every(file=>officeImportable(file.name)))current.current.onImport(files);else if(!current.current.onOfficeImport(files))awaitingDialog=false;}
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
  const office=(files:File[])=>{const accepted=files.filter(file=>officeImportable(file.name)),legacy=files.filter(file=>LEGACY_OFFICE.test(file.name)).length,other=files.length-accepted.length-legacy;
   const skipped=[legacy?`${legacy} 个旧版 .doc / .xls / .ppt 文件未导入，请在 Office 中另存为 .docx / .xlsx / .pptx`:'',other?`${other} 个非办公文件未导入，图片和附件请拖入正文或白板`:''].filter(Boolean).join('；');
   if(accepted.length&&!current.current.onOfficeImport(accepted))return;if(skipped)current.current.notify(skipped+'。');};
  const attach=(files:File[])=>{
   const host=document.querySelector('.workbench-pane.active-pane .whiteboard')??document.querySelector('.workbench-pane.active-pane .markdown-document');
   const request=new CustomEvent('mydoc:insert-files',{detail:files,cancelable:true});host?.dispatchEvent(request);
   if(request.defaultPrevented)return;
   // Without an open document, an all-Office drop (including PDF) imports as files.
   if(!host&&files.every(file=>officeImportable(file.name)))current.current.onOfficeImport(files);
   else current.current.notify('请先打开文档，再把图片或附件拖入正文或白板。');
  };
  const capture=(event:DragEvent)=>{if(!event.dataTransfer?.types.includes('Files'))return;reset();
   const target=event.target instanceof Element?event.target:null;
   // Import dialogs and panels mark their own drop targets.
   if(target?.closest('[data-file-drop]'))return;
   if(modalOpen()){event.preventDefault();event.stopPropagation();// An open import dialog already shows where to drop; stay quiet on a miss.
    if(![...document.querySelectorAll(MODAL)].some(modal=>modal.matches('[data-file-drop]')||modal.querySelector('[data-file-drop]')))current.current.notify('请先完成或关闭当前弹窗，再拖入文件。');return;}
   if(!current.current.ready||current.current.busy){event.preventDefault();event.stopPropagation();current.current.notify('正在处理其他操作，请稍后重新拖入文件。');return;}
   if([...event.dataTransfer.items].some(item=>item.kind==='file'&&item.webkitGetAsEntry?.()?.isDirectory)){event.preventDefault();event.stopPropagation();current.current.notify('拖入文件夹请使用导入窗口的“选择文件夹”，或先压缩为 ZIP。');return;}
   const files=[...event.dataTransfer.files],insideEditor=!!target?.closest('.ProseMirror,.whiteboard-canvas');
   if(importableDrop(files,insideEditor)){event.preventDefault();event.stopPropagation();current.current.onImport(files);return;}
   if(officeDrop(files,insideEditor)){event.preventDefault();event.stopPropagation();office(files);return;}
   // A single MyDoc JSON file (backup, draft, template) opens its restore/import flow.
   if(!insideEditor&&files.length===1&&/\.json$/i.test(files[0].name)){event.preventDefault();event.stopPropagation();const file=files[0];
    void file.text().then(text=>{const kind=file.size<=50*1024*1024?jsonImportKind(text):undefined;if(kind)current.current.onJsonImport(file,kind);else attach(files);}).catch(()=>current.current.notify('无法读取拖入的文件，请重试。'));}
  };
  const drop=(event:DragEvent)=>{if(!event.dataTransfer?.types.includes('Files')||event.defaultPrevented)return;event.preventDefault();const files=[...event.dataTransfer.files];if(!files.length)return;attach(files);};
  document.addEventListener('dragenter',enter);document.addEventListener('dragleave',leave);document.addEventListener('dragover',over);document.addEventListener('drop',capture,true);document.addEventListener('drop',drop);window.addEventListener('dragend',reset);window.addEventListener('blur',reset);
  return()=>{document.removeEventListener('dragenter',enter);document.removeEventListener('dragleave',leave);document.removeEventListener('dragover',over);document.removeEventListener('drop',capture,true);document.removeEventListener('drop',drop);window.removeEventListener('dragend',reset);window.removeEventListener('blur',reset);};
 },[]);
 return dragging;
}
