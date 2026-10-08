import { useEffect, useRef } from 'react';
import { ChevronUp, ChevronDown, X } from 'lucide-react';
import type { EditorView } from 'prosemirror-view';
import { MAX_FIND_QUERY, updateDocumentFind, type DocumentFindState } from './document-find';
export function DocumentFindBar({state,view,onClose}:{state:DocumentFindState;view:()=>EditorView|null;onClose:()=>void}){
  const input=useRef<HTMLInputElement>(null);
  useEffect(()=>{input.current?.focus({preventScroll:true});input.current?.select();input.current?.closest('.document-find-bar')?.scrollIntoView({block:'nearest',inline:'nearest'});},[]);
  const move=(direction:1|-1)=>{const editor=view();if(editor)updateDocumentFind(editor,{direction});};
  const count=state.query?state.matches.length?`${state.index+1} / ${state.matches.length}${state.truncated?'+':''}`:'无匹配':'输入查找内容';
  return <div className="document-find-bar" role="search" aria-label="页内查找" onKeyDown={event=>{if(event.nativeEvent.isComposing||event.keyCode===229)return;if(event.key==='Escape'){event.preventDefault();event.stopPropagation();onClose();}else if(event.key==='Enter'){event.preventDefault();move(event.shiftKey?-1:1);}}}>
    <label className="document-find-input"><span className="sr-only">查找正文</span><input ref={input} type="search" aria-label="查找正文" placeholder="查找正文…" value={state.query} maxLength={MAX_FIND_QUERY} onChange={event=>{const editor=view();if(editor)updateDocumentFind(editor,{query:event.target.value});}}/></label>
    <span className="document-find-count" role="status" aria-live="polite">{count}</span>
    <button type="button" aria-label="上一个匹配" title="上一个匹配（Shift + Enter）" disabled={!state.matches.length} onClick={()=>move(-1)}><ChevronUp size={16}/></button>
    <button type="button" aria-label="下一个匹配" title="下一个匹配（Enter）" disabled={!state.matches.length} onClick={()=>move(1)}><ChevronDown size={16}/></button>
    <label className="document-find-case"><input type="checkbox" checked={state.caseSensitive} onChange={event=>{const editor=view();if(editor)updateDocumentFind(editor,{caseSensitive:event.target.checked});}}/>区分大小写</label>
    <button type="button" aria-label="关闭页内查找" onClick={onClose}><X size={16}/></button>
  </div>;
}
