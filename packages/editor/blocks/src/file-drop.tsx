import {useRef,useState,type DragEvent} from 'react';
interface FileDropOptions {disabled?:boolean;multiple?:boolean;onFiles:(files:File[])=>void;onError:(message:string)=>void;directoryMessage?:string;}
/**
 * Drag-and-drop target for an import surface. The data-file-drop marker tells the
 * app-wide intake (file-intake.ts) to leave the drop to this element.
 */
export function useFileDrop({disabled=false,multiple=false,onFiles,onError,directoryMessage='请拖入文件；文件夹可先压缩为 ZIP。'}:FileDropOptions){
 const [dragging,setDragging]=useState(false);const depth=useRef(0);
 const files=(event:DragEvent)=>event.dataTransfer.types.includes('Files');
 const props={
  'data-file-drop':'',
  onDragEnter(event:DragEvent<HTMLElement>){if(!files(event))return;event.preventDefault();if(!disabled){depth.current++;setDragging(true);}},
  onDragOver(event:DragEvent<HTMLElement>){if(!files(event))return;event.preventDefault();event.stopPropagation();event.dataTransfer.dropEffect=disabled?'none':'copy';},
  onDragLeave(){depth.current=Math.max(0,depth.current-1);if(!depth.current)setDragging(false);},
  onDrop(event:DragEvent<HTMLElement>){if(!files(event))return;event.preventDefault();event.stopPropagation();depth.current=0;setDragging(false);if(disabled)return;
   if([...event.dataTransfer.items].some(item=>item.kind==='file'&&item.webkitGetAsEntry?.()?.isDirectory)){onError(directoryMessage);return;}
   const dropped=[...event.dataTransfer.files];
   if(!dropped.length)onError('未能读取拖入的文件，请使用选择文件按钮重试。');
   else if(!multiple&&dropped.length>1)onError('一次只能拖入一个文件');
   else onFiles(dropped);
  },
 };
 return {dragging,props};
}
export type FileDropProps=ReturnType<typeof useFileDrop>['props'];
export function FileDropHint({dragging,disabled,label='导入文件拖放区',hint}:{dragging:boolean;disabled?:boolean;label?:string;hint:string}){
 return <div role="region" aria-label={label} aria-disabled={disabled} className={`import-drop-zone${dragging?' dragging':''}`}><strong>{dragging?'松开鼠标，读取文件':'拖动文件到这里导入'}</strong><span>{hint}</span></div>;
}
