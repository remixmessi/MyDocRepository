import {useLayoutEffect,useMemo,useState,type RefObject} from 'react';
import type {Y} from '@mydoc/crdt';
import {captureTextMatch,findTextRanges,replaceTextMatches,resolveTextMatch,type TextMatch} from './office-text-find';

export function TextFindPanel({text,input,queryInput,readOnly,composing,assertWritable,boundary,onClose}:{text:Y.Text;input:RefObject<HTMLTextAreaElement>;queryInput:RefObject<HTMLTextAreaElement>;readOnly:boolean;composing:boolean;assertWritable:()=>void;boundary:()=>void;onClose:()=>void}){
 const [query,setQuery]=useState(''),[replacement,setReplacement]=useState(''),[active,setActive]=useState<TextMatch>(),[message,setMessage]=useState('');
 const source=text.toString(),result=useMemo(()=>findTextRanges(source,query),[source,query]);
 const resolved=active&&resolveTextMatch(text,active),index=resolved?result.ranges.findIndex(range=>range.start===resolved.start&&range.end===resolved.end):-1;
 const selected=index<0?undefined:resolved;
 useLayoutEffect(()=>{
  const element=input.current;
  if(!composing&&element&&selected&&document.activeElement===element)element.setSelectionRange(selected.start,selected.end);
 },[source,selected?.start,selected?.end,composing,input]);
 const locate=(direction:1|-1)=>{
  if(composing||!result.ranges.length)return;
  const next=index<0?(direction===1?0:result.ranges.length-1):(index+direction+result.ranges.length)%result.ranges.length;
  const range=result.ranges[next];setActive(captureTextMatch(text,range));setMessage('');
  input.current?.focus();input.current?.setSelectionRange(range.start,range.end);
 };
 const replace=(all:boolean)=>{
  if(readOnly||composing||!query||!all&&!selected)return;
  try{assertWritable();boundary();const count=replaceTextMatches(text,query,replacement,all?undefined:active);setActive(undefined);setMessage(`已替换 ${count} 处`);}catch(error){setMessage(error instanceof Error?error.message:'替换失败');}finally{boundary();}
 };
 return <div className="text-find-panel" role="group" aria-label="文本查找替换" onKeyDown={event=>{if(event.key==='Escape'&&!event.nativeEvent.isComposing){event.preventDefault();event.stopPropagation();onClose();input.current?.focus();}}}>
  <label>查找<textarea ref={queryInput} autoFocus rows={1} aria-label="查找文本" maxLength={10000} value={query} onChange={event=>{setQuery(event.target.value);setActive(undefined);setMessage('');}}/></label>
  <div className="text-find-actions"><button disabled={composing||!result.ranges.length} onClick={()=>locate(-1)}>上一处</button><button disabled={composing||!result.ranges.length} onClick={()=>locate(1)}>下一处</button></div>
  <label>替换为<textarea rows={1} aria-label="替换文本" maxLength={10000} value={replacement} disabled={readOnly||composing} onChange={event=>setReplacement(event.target.value)}/></label>
  <div className="text-find-actions"><button disabled={readOnly||composing||!selected} onClick={()=>replace(false)}>替换当前</button><button disabled={readOnly||composing||!result.ranges.length||result.truncated} onClick={()=>replace(true)}>全部替换</button><button onClick={onClose}>关闭查找</button></div>
  <p role="status">{!query?'区分大小写的字面查找，支持多行文本。':result.truncated?'匹配超过 1 万处，请缩小查找范围。':result.ranges.length?`${index<0?'未定位':`第 ${index+1} 处`} / 共 ${result.ranges.length} 处匹配`:'没有匹配内容'}{composing?' · 输入法正在输入，暂不可定位或替换':''}{message&&` · ${message}`}{readOnly?' · 只读，可查找':''}</p>
 </div>;
}
