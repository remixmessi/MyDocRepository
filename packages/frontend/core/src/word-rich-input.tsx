import {wordTextIdentity} from './office-word-flow';
import {useEffect,useLayoutEffect,useRef} from 'react';
import {Y} from '@mydoc/crdt';
import {validWordLink} from '../../../common/nbstore/src/office-word-link';
import {WORD_HIGHLIGHTS,validWordHighlight,validWordFontFamily,validWordVertAlign} from '../../../common/nbstore/src/office-word-font';
import {validWordRevisionMark,validWordRunFormatMark,wordRevisionColor} from '../../../common/nbstore/src/office-word-revision';
import {WORD_NOTE_CHAR,validWordNoteId} from '../../../common/nbstore/src/office-word-reference';

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
 // A note reference is not editable: the caret goes just before or after it.
 const reference=(node:Node)=>node.parentElement?.closest<HTMLElement>('.word-note-ref,.word-note-stray');
 const around=(span:HTMLElement,after:boolean):[Node,number]=>{const parent=span.parentNode!;return [parent,[...parent.childNodes].indexOf(span)+(after?1:0)];};
 const point=(offset:number):[Node,number]=>{for(const [i,node]of nodes.entries()){const span=reference(node);if(span&&offset<=node.length){if(offset===0)return around(span,false);if(offset===node.length&&!nodes[i+1])return around(span,true);if(offset===node.length){offset-=node.length;continue;}}if(offset<node.length||offset===node.length&&!separator(node))return[node,offset];if(offset===node.length){const next=nodes[i+1];return next?[next,0]:[element,element.childNodes.length-(element.lastChild?.nodeName==='BR'?1:0)];}offset-=node.length;}return[nodes[nodes.length-1],nodes[nodes.length-1].length];};
 const range=document.createRange();range.setStart(...point(start));range.setEnd(...point(end));const selection=window.getSelection();selection?.removeAllRanges();selection?.addRange(range);
}
function render(element:HTMLElement|DocumentFragment,text:Y.Text|undefined,parts:{insert?:unknown;attributes?:Record<string,any>}[]=text!.toDelta()){
 const fragment=document.createDocumentFragment();
 for(const part of parts){
  // A note reference character renders as its own non-editable superscript; other text never carries a reference.
  if(String(part.insert??'').includes(WORD_NOTE_CHAR)){const value=String(part.insert);let rest='';const flush=()=>{if(rest){render(fragment,undefined,[{insert:rest,attributes:{...part.attributes,note:undefined}}]);rest='';}};for(const char of value){if(char!==WORD_NOTE_CHAR){rest+=char;continue;}flush();const ref=document.createElement('span'),id=part.attributes?.note;ref.className=validWordNoteId(id)?'word-note-ref':'word-note-stray';if(validWordNoteId(id))ref.dataset.note=id;ref.contentEditable='false';const glyph=document.createElement('span');glyph.className='word-note-glyph';glyph.textContent=WORD_NOTE_CHAR;ref.append(glyph);for(const kind of ['ins','del'] as const){const mark=part.attributes?.[kind];if(validWordRevisionMark(mark)){ref.dataset.revision=kind;ref.style.setProperty('--revision-color',wordRevisionColor(mark.author));}}fragment.append(ref);}flush();continue;}
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
  // Tracked changes: insertions underlined and deletions struck through, in the reviewer's colour.
  for(const kind of ['ins','del'] as const){const mark=marks[kind];if(!validWordRevisionMark(mark))continue;span.dataset.revision=kind;span.dataset.revisionId=mark.id;span.style.setProperty('--revision-color',wordRevisionColor(mark.author));span.title=`${mark.name} ${kind==='ins'?'插入':'删除'}：${new Date(mark.date).toLocaleString()}`;}
  if(validWordRunFormatMark(marks.fmt)&&!span.dataset.revision){span.dataset.formatRevision='true';span.style.setProperty('--revision-color',wordRevisionColor(marks.fmt.author));span.title=`${marks.fmt.name} 设置格式：${new Date(marks.fmt.date).toLocaleString()}`;}
  fragment.append(span);
 }
 if(!text){element.append(fragment);return;}
 // A trailing line or paragraph break needs a <br> so the empty last line renders and can hold the caret.
 if(!fragment.childNodes.length||/[\n\u2029]$/.test(text.toString()))fragment.append(document.createElement('br'));
 // Note numbers are set by the editor across all paragraphs; keep them when only this paragraph re-renders.
 const labels=new Map([...(element as HTMLElement).querySelectorAll<HTMLElement>('.word-note-ref')].map(node=>[node.dataset.note,node.dataset.label]));for(const node of fragment.querySelectorAll<HTMLElement>('.word-note-ref')){const label=labels.get(node.dataset.note);if(label)node.dataset.label=label;}
 (element as HTMLElement).replaceChildren(fragment);
}
export function WordRichInput({text,label,readOnly,onCommit,onCompositionStart,onCompositionEnd,onSelection,onFocus,onFormat,onEnter,onBoundary,onNavigate,paragraphText}:{paragraphText?:string;text:Y.Text;label:string;readOnly:boolean;onCommit:(value:string,direction:'backward'|'forward')=>number|void;onCompositionStart:()=>void;onCompositionEnd:(value:string)=>boolean;onSelection:(start:number,end:number)=>void;onFocus:()=>void;onFormat:(mark:'bold'|'italic'|'underline')=>void;onEnter?:(start:number,end:number)=>void;onBoundary?:(direction:'backward'|'forward')=>boolean;onNavigate?:(direction:'backward'|'forward')=>boolean}){
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
 const insert=(raw:string)=>{const value=raw.replaceAll(WORD_NOTE_CHAR,'');if(!value&&raw)return;lastInput.current='insertText';const element=ref.current!,selection=window.getSelection();if(!selection?.rangeCount)return;const range=selection.getRangeAt(0);if(!element.contains(range.commonAncestorContainer))return;range.deleteContents();const node=document.createTextNode(value);range.insertNode(node);range.setStartAfter(node);range.collapse(true);selection.removeAllRanges();selection.addRange(range);commit(element.textContent??'');};
 // Which way the last deletion went, so a tracked deletion leaves the caret on the right side of the struck text.
 const lastInput=useRef('');
 const commit=(value:string)=>{editing.current=true;let caret:number|void=undefined;try{caret=onCommit(value,lastInput.current==='deleteContentForward'?'forward':'backward');}finally{editing.current=false;const element=ref.current;
  // A tracked edit keeps deleted text, so the DOM no longer equals the input; re-render and place the returned caret.
  if(element&&typeof caret==='number'){render(element,text);restoreWordCaret(element,Math.min(caret,text.length),Math.min(caret,text.length));}
  else if(element&&element.textContent===text.toString()){const range=wordDomSelection(element);render(element,text);if(range)restoreWordCaret(element,range.start,range.end);}}};
 useEffect(()=>{
  const element=ref.current!;
  const beforeInput=(event:InputEvent)=>{
   lastInput.current=event.inputType;
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
  // Chrome's End/Home stop short of a non-editable note reference at the end/start of the line; on that line, go past it.
  if((event.key==='End'||event.key==='Home')&&!event.shiftKey&&!event.ctrlKey&&!event.metaKey&&!event.altKey){const value=text.toString(),end=event.key==='End',edge=end?value.endsWith(WORD_NOTE_CHAR):value.startsWith(WORD_NOTE_CHAR),refs=event.currentTarget.querySelectorAll<HTMLElement>('.word-note-ref,.word-note-stray'),ref=end?refs[refs.length-1]:refs[0],dom=window.getSelection();
   if(edge&&ref&&dom?.rangeCount){const caret=dom.getRangeAt(0).getClientRects()[0]??dom.getRangeAt(0).getBoundingClientRect(),box=ref.getBoundingClientRect();if(Math.abs((caret.top+caret.bottom)/2-(box.top+box.bottom)/2)<=Math.max(box.height,caret.height)){event.preventDefault();restoreWordCaret(event.currentTarget,end?value.length:0,end?value.length:0);selection();return;}}}
  const range=wordDomSelection(event.currentTarget);
  if(range&&range.start===range.end&&!event.metaKey&&!event.ctrlKey&&!event.altKey){
   const backward=range.start===0,forward=range.end===text.length;
   if(event.key==='Backspace'&&backward&&onBoundary?.('backward')||event.key==='Delete'&&forward&&onBoundary?.('forward')){event.preventDefault();return;}
   if(!event.shiftKey&&((event.key==='ArrowLeft'||event.key==='ArrowUp')&&backward||(event.key==='ArrowRight'||event.key==='ArrowDown')&&forward)&&onNavigate?.(backward&&(event.key==='ArrowLeft'||event.key==='ArrowUp')?'backward':'forward')){event.preventDefault();return;}
  }
  if(event.key==='Enter'){event.preventDefault();const range=wordDomSelection(event.currentTarget);if(!event.shiftKey&&onEnter&&range)onEnter(range.start,range.end);else insert(!event.shiftKey&&paragraphText?paragraphText:'\n');}
 }} onDrop={event=>{event.preventDefault();if(!readOnly)insert(event.dataTransfer.getData('text/plain'));}} onPaste={event=>{if(readOnly)return;event.preventDefault();insert(event.clipboardData.getData('text/plain'));}}/>
}
