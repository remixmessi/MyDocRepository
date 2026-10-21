import {useEffect,useRef,useState,type RefObject} from 'react';
import {Y} from '@mydoc/crdt';
import {readWordFlowDom,restoreWordFlowDom,wordDomPoint} from './word-flow-dom';
import {formatWordFlow,replaceWordFlow,wordFlowPlainText,type WordFlowSelection} from './word-flow-selection';
import {wordTextIdentity} from './office-word-flow';
import type {WordMark,WordFontKey} from './office-word-format';
import type {WordBlock} from '../../../common/nbstore/src/office-word-table';
export function useWordFlowSelection(root:RefObject<HTMLElement>,doc:Y.Doc,readOnly:boolean,assertWritable:()=>void,focus:(text:Y.Text,offset?:number)=>void,boundary:()=>void){
 const range=useRef<WordFlowSelection>(),mutating=useRef(false),[active,setActive]=useState(false),[error,setError]=useState('');
 const clear=()=>{range.current=undefined;setActive(false);};
 const replace=(value:string)=>{if(readOnly)return;try{assertWritable();const selected=range.current;if(!selected)return;mutating.current=true;boundary();const result=replaceWordFlow(doc,selected,value);clear();setError('');focus(result.text,result.offset);}catch(error){setError(error instanceof Error?error.message:'正文编辑失败，请保留选区内容');}finally{boundary();mutating.current=false;}};
 const format=(key:WordMark|WordFontKey|'clear',value?:string|number)=>{if(!range.current)return false;if(readOnly)return true;try{assertWritable();boundary();formatWordFlow(doc,range.current,key,value);boundary();setError('');}catch(error){setError(error instanceof Error?error.message:'选区格式修改失败');}return true;};
 useEffect(()=>{
  const element=root.current;if(!element)return;let restore:WordFlowSelection|undefined;
  const select=()=>{if(mutating.current)return;const dom=window.getSelection();if(!dom?.rangeCount)return;const selected=dom.getRangeAt(0);if(!element.contains(selected.startContainer)||!element.contains(selected.endContainer))return;const flow=readWordFlowDom(element,doc);if(flow){range.current=flow;setActive(true);}else if((selected.startContainer instanceof HTMLElement?selected.startContainer:selected.startContainer.parentElement)?.closest('.word-paper'))clear();};
  const before=()=>{if(mutating.current)return;restore=readWordFlowDom(element,doc);if(restore){range.current=restore;setActive(true);}};
  const after=()=>{if(mutating.current||!restore)return;try{restoreWordFlowDom(element,doc,restore);}catch{clear();}restore=undefined;};
  const input=(event:InputEvent)=>{select();if(!range.current||event.defaultPrevented)return;if(readOnly){event.preventDefault();return;}
   const type=event.inputType;if(type.startsWith('delete')||['insertText','insertReplacementText','insertParagraph','insertLineBreak'].includes(type)){event.preventDefault();event.stopImmediatePropagation();replace(type.startsWith('delete')?'':type==='insertParagraph'||type==='insertLineBreak'?'\n':event.data??'');}
  };
  const key=(event:KeyboardEvent)=>{
   if(event.isComposing)return;select();const dom=window.getSelection();
   if(event.shiftKey&&dom?.focusNode&&['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)){
    const input=(dom.focusNode instanceof HTMLElement?dom.focusNode:dom.focusNode.parentElement)?.closest<HTMLElement>('.word-rich-input');
    if(input&&element.contains(input)){const blocks=doc.getArray<WordBlock>('paragraphs').toArray(),index=blocks.findIndex(text=>text instanceof Y.Text&&wordTextIdentity(text)===input.dataset.wordText),text=blocks[index];if(text instanceof Y.Text){const offsetRange=document.createRange();offsetRange.selectNodeContents(input);offsetRange.setEnd(dom.focusNode,dom.focusOffset);const offset=offsetRange.toString().length,backward=event.key==='ArrowLeft'||event.key==='ArrowUp',next=blocks[index+(backward?-1:1)];if((backward&&offset===0||!backward&&offset===text.length)&&next instanceof Y.Text){const target=element.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(next)}"]`);if(target){event.preventDefault();event.stopPropagation();dom.extend(...wordDomPoint(target,backward?next.length:0));select();return;}}}}
   }
   if(!range.current)return;
   if((event.ctrlKey||event.metaKey)&&['b','i','u'].includes(event.key.toLowerCase())){event.preventDefault();event.stopPropagation();format(event.key.toLowerCase()==='b'?'bold':event.key.toLowerCase()==='i'?'italic':'underline');return;}
   if(['Backspace','Delete','Enter'].includes(event.key)||!event.ctrlKey&&!event.metaKey&&!event.altKey&&event.key.length===1){event.preventDefault();event.stopPropagation();replace(event.key==='Enter'?'\n':event.key==='Backspace'||event.key==='Delete'?'':event.key);}
  };
  const paste=(event:ClipboardEvent)=>{select();if(!range.current)return;event.preventDefault();event.stopPropagation();if(!readOnly)replace(event.clipboardData?.getData('text/plain')??'');};
  const copy=(event:ClipboardEvent)=>{select();if(!range.current)return;try{const value=wordFlowPlainText(doc,range.current);event.preventDefault();event.stopPropagation();event.clipboardData?.setData('text/plain',value);if(event.type==='cut'&&!readOnly)replace('');}catch(error){event.preventDefault();setError(error instanceof Error?error.message:'无法复制选区');}};
  const compose=()=>{select();if(range.current&&!readOnly)replace('');};
  document.addEventListener('selectionchange',select);element.addEventListener('beforeinput',input,true);element.addEventListener('keydown',key,true);element.addEventListener('paste',paste,true);element.addEventListener('copy',copy,true);element.addEventListener('cut',copy,true);element.addEventListener('compositionstart',compose,true);doc.on('beforeTransaction',before);doc.on('afterTransaction',after);
  return()=>{document.removeEventListener('selectionchange',select);element.removeEventListener('beforeinput',input,true);element.removeEventListener('keydown',key,true);element.removeEventListener('paste',paste,true);element.removeEventListener('copy',copy,true);element.removeEventListener('cut',copy,true);element.removeEventListener('compositionstart',compose,true);doc.off('beforeTransaction',before);doc.off('afterTransaction',after);};
 });
 return {active,error,format};
}
