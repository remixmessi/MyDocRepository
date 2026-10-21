import {wordTextIdentity} from './office-word-flow';
import {useEffect,useLayoutEffect,useRef} from 'react';
import {Y} from '@mydoc/crdt';
import {validWordLink} from '../../../common/nbstore/src/office-word-link';
import {WORD_HIGHLIGHTS,validWordHighlight,validWordFontFamily,validWordVertAlign} from '../../../common/nbstore/src/office-word-font';

/** DOM offsets are UTF-16 offsets, matching Y.Text and native browser selections. */
export function wordDomSelection(element:HTMLElement){
 const selection=window.getSelection();if(!selection?.rangeCount)return;
 const range=selection.getRangeAt(0);if(!element.contains(range.startContainer)||!element.contains(range.endContainer))return;
 const before=range.cloneRange();before.selectNodeContents(element);before.setEnd(range.startContainer,range.startOffset);
 return {start:before.toString().length,end:before.toString().length+range.toString().length};
}
export function restoreWordCaret(element:HTMLElement,start:number,end:number){
 const walker=document.createTreeWalker(element,NodeFilter.SHOW_TEXT),nodes:Text[]=[];let node:Node|null;
 while(node=walker.nextNode())nodes.push(node as Text);
 if(!nodes.length){const range=document.createRange();range.selectNodeContents(element);range.collapse(true);const selection=window.getSelection();selection?.removeAllRanges();selection?.addRange(range);return;}
 // Never leave the caret inside a hidden cell paragraph separator; place it at the start of the next paragraph.
 const separator=(node:Node)=>!!node.parentElement?.classList.contains('word-cell-paragraph-break');
 const point=(offset:number):[Node,number]=>{for(const [i,node]of nodes.entries()){if(offset<node.length||offset===node.length&&!separator(node))return[node,offset];if(offset===node.length){const next=nodes[i+1];return next?[next,0]:[element,element.childNodes.length-(element.lastChild?.nodeName==='BR'?1:0)];}offset-=node.length;}return[nodes[nodes.length-1],nodes[nodes.length-1].length];};
 const range=document.createRange();range.setStart(...point(start));range.setEnd(...point(end));const selection=window.getSelection();selection?.removeAllRanges();selection?.addRange(range);
}
function render(element:HTMLElement,text:Y.Text){
 const fragment=document.createDocumentFragment();
 for(const part of text.toDelta()){
  const marks=part.attributes??{},span=document.createElement(validWordLink(marks.href)?'a':'span');
  // Cell paragraph separators keep their character for 1:1 offsets but render as a block break.
  const value=String(part.insert??''),pieces=value.split(/(\u2029)/).filter(Boolean);if(pieces.length>1||value==='\u2029'){for(const piece of pieces){const inner=document.createElement('span');inner.textContent=piece;if(piece==='\u2029')inner.className='word-cell-paragraph-break';span.append(inner);}}else span.textContent=value;if(span instanceof HTMLAnchorElement){span.href=marks.href;span.title=marks.href;span.target='_blank';span.rel='noopener noreferrer';}
  if(typeof marks.fontSize==='number')span.style.fontSize=marks.fontSize+'pt';
  if(typeof marks.color==='string'&&marks.color!=='auto')span.style.color=marks.color;
  if(marks.bold!==undefined)span.style.fontWeight=marks.bold?'bold':'normal';
  if(marks.italic!==undefined)span.style.fontStyle=marks.italic?'italic':'normal';
  if(marks.underline!==undefined||marks.strike!==undefined)span.style.textDecoration=[marks.underline?'underline':'',marks.strike?'line-through':''].filter(Boolean).join(' ')||'none';
  if(validWordHighlight(marks.highlight))span.style.backgroundColor=WORD_HIGHLIGHTS[marks.highlight];
  if(validWordFontFamily(marks.fontFamily))span.style.fontFamily=`"${marks.fontFamily}"`;
  if(validWordVertAlign(marks.vertAlign)){span.style.verticalAlign=marks.vertAlign==='superscript'?'super':'sub';span.style.fontSize=typeof marks.fontSize==='number'?marks.fontSize*0.65+'pt':'0.65em';}
  fragment.append(span);
 }
 // A trailing line or paragraph break needs a <br> so the empty last line renders and can hold the caret.
 if(!fragment.childNodes.length||/[\n\u2029]$/.test(text.toString()))fragment.append(document.createElement('br'));
 element.replaceChildren(fragment);
}
export function WordRichInput({text,label,readOnly,onCommit,onCompositionStart,onCompositionEnd,onSelection,onFocus,onFormat,onEnter,onBoundary,onNavigate,paragraphText}:{paragraphText?:string;text:Y.Text;label:string;readOnly:boolean;onCommit:(value:string)=>void;onCompositionStart:()=>void;onCompositionEnd:(value:string)=>boolean;onSelection:(start:number,end:number)=>void;onFocus:()=>void;onFormat:(mark:'bold'|'italic'|'underline')=>void;onEnter?:(start:number,end:number)=>void;onBoundary?:(direction:'backward'|'forward')=>boolean;onNavigate?:(direction:'backward'|'forward')=>boolean}){
 const ref=useRef<HTMLDivElement>(null),composing=useRef(false),editing=useRef(false);
 const onSelect=useRef(onSelection);onSelect.current=onSelection;
 const selection=()=>{const element=ref.current;if(!element)return;const range=wordDomSelection(element);if(range)onSelect.current(range.start,range.end);};
 useLayoutEffect(()=>{
  const element=ref.current!;
  let saved:{start:Y.RelativePosition;end:Y.RelativePosition}|undefined;
  const before=()=>{const range=wordDomSelection(element);saved=range?{start:Y.createRelativePositionFromTypeIndex(text,range.start),end:Y.createRelativePositionFromTypeIndex(text,range.end)}:undefined;};
  const update=()=>{
   if(composing.current||editing.current)return;
   const start=saved?.start,end=saved?.end;
   render(element,text);
   if(start&&end){const left=Y.createAbsolutePositionFromRelativePosition(start,text.doc!),right=Y.createAbsolutePositionFromRelativePosition(end,text.doc!);if(left&&right)restoreWordCaret(element,left.index,right.index);}
  };
  update();text.observe(update);text.doc!.on('beforeTransaction',before);return()=>{text.unobserve(update);text.doc?.off('beforeTransaction',before);};
 },[text]);
 useEffect(()=>{document.addEventListener('selectionchange',selection);return()=>document.removeEventListener('selectionchange',selection);});
 const insert=(value:string)=>{const element=ref.current!,selection=window.getSelection();if(!selection?.rangeCount)return;const range=selection.getRangeAt(0);if(!element.contains(range.commonAncestorContainer))return;range.deleteContents();const node=document.createTextNode(value);range.insertNode(node);range.setStartAfter(node);range.collapse(true);selection.removeAllRanges();selection.addRange(range);commit(element.textContent??'');};
 const commit=(value:string)=>{editing.current=true;try{onCommit(value);}finally{editing.current=false;const element=ref.current;if(element&&element.textContent===text.toString()){const range=wordDomSelection(element);render(element,text);if(range)restoreWordCaret(element,range.start,range.end);}}};
 useEffect(()=>{
  const element=ref.current!;
  const beforeInput=(event:InputEvent)=>{
   if(readOnly||composing.current||event.isComposing||event.defaultPrevented)return;
   const range=wordDomSelection(element);
   if(['insertParagraph','insertLineBreak'].includes(event.inputType)){event.preventDefault();if(event.inputType==='insertParagraph'&&onEnter&&range)onEnter(range.start,range.end);else insert(event.inputType==='insertParagraph'&&paragraphText?paragraphText:'\n');}
   else if(range&&range.start===range.end&&((event.inputType==='deleteContentBackward'&&range.start===0)||(event.inputType==='deleteContentForward'&&range.end===text.length))){if(onBoundary?.(event.inputType==='deleteContentBackward'?'backward':'forward'))event.preventDefault();}
  };
  element.addEventListener('beforeinput',beforeInput);return()=>element.removeEventListener('beforeinput',beforeInput);
 });
 return <div ref={ref} className="word-rich-input" data-word-text={wordTextIdentity(text)} role="textbox" tabIndex={0} aria-label={label} aria-multiline="true" aria-readonly={readOnly} contentEditable={!readOnly} suppressContentEditableWarning spellCheck onFocus={onFocus}
 onInput={event=>{if(!composing.current)commit(event.currentTarget.textContent??'');}}
 onCompositionStart={()=>{composing.current=true;onCompositionStart();}}
 onCompositionEnd={event=>{composing.current=false;editing.current=true;let applied=false;try{applied=onCompositionEnd(event.currentTarget.textContent??'');}finally{editing.current=false;const element=ref.current;if(element&&applied){const range=wordDomSelection(element);render(element,text);if(range)restoreWordCaret(element,Math.min(range.start,text.length),Math.min(range.end,text.length));}}}}
 onKeyDown={event=>{
  if(readOnly||composing.current)return;
  if((event.ctrlKey||event.metaKey)&&['b','i','u'].includes(event.key.toLowerCase())){event.preventDefault();selection();onFormat(event.key.toLowerCase()==='b'?'bold':event.key.toLowerCase()==='i'?'italic':'underline');}
  const range=wordDomSelection(event.currentTarget);
  if(range&&range.start===range.end&&!event.metaKey&&!event.ctrlKey&&!event.altKey){
   const backward=range.start===0,forward=range.end===text.length;
   if(event.key==='Backspace'&&backward&&onBoundary?.('backward')||event.key==='Delete'&&forward&&onBoundary?.('forward')){event.preventDefault();return;}
   if(!event.shiftKey&&((event.key==='ArrowLeft'||event.key==='ArrowUp')&&backward||(event.key==='ArrowRight'||event.key==='ArrowDown')&&forward)&&onNavigate?.(backward&&(event.key==='ArrowLeft'||event.key==='ArrowUp')?'backward':'forward')){event.preventDefault();return;}
  }
  if(event.key==='Enter'){event.preventDefault();const range=wordDomSelection(event.currentTarget);if(!event.shiftKey&&onEnter&&range)onEnter(range.start,range.end);else insert(!event.shiftKey&&paragraphText?paragraphText:'\n');}
 }} onDrop={event=>{event.preventDefault();if(!readOnly)insert(event.dataTransfer.getData('text/plain'));}} onPaste={event=>{if(readOnly)return;event.preventDefault();insert(event.clipboardData.getData('text/plain'));}}/>
}
