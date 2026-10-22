import {useRef,useState} from 'react';
import type {Y} from '@mydoc/crdt';
import {Modal} from './dialogs';
import {getWordHeaderFooter,setWordHeaderFooter} from './office-word-page';
import {EMPTY_WORD_LINE,WORD_PAGE_TOKEN,WORD_PAGES_TOKEN,emptyWordLine,type WordHeaderFooter,type WordHeaderLine} from '../../../common/nbstore/src/office-word-header';

type Part='header'|'footer';
const SLOTS=[['left','左'],['center','中'],['right','右']] as const;
const blank=():WordHeaderFooter=>({header:{...EMPTY_WORD_LINE},footer:{...EMPTY_WORD_LINE}});
/** 插入 → 页眉和页脚: left/centre/right text per line, page-number fields and a different first page, like Word. */
export function WordHeaderDialog({doc,readOnly,focus='header',onClose}:{doc:Y.Doc;readOnly:boolean;focus?:Part;onClose:()=>void}){
 const [base,setBase]=useState(()=>getWordHeaderFooter(doc)),[draft,setDraft]=useState<WordHeaderFooter>(()=>base??blank()),[error,setError]=useState('');
 const last=useRef<{input:HTMLInputElement;set:(text:string)=>void}>();
 const reload=()=>{const value=getWordHeaderFooter(doc);setBase(value);setDraft(value??blank());setError('');};
 const line=(scope:'main'|'first',part:Part):WordHeaderLine=>scope==='first'?draft.first![part]:draft[part];
 const update=(scope:'main'|'first',part:Part,slot:keyof WordHeaderLine,text:string)=>setDraft(value=>scope==='first'?{...value,first:{...value.first!,[part]:{...value.first![part],[slot]:text}}}:{...value,[part]:{...value[part],[slot]:text}});
 // Field buttons insert at the caret of the last focused box, as Word's 页码 button does.
 const insert=(token:string)=>{const target=last.current;if(!target)return;const {input}=target,start=input.selectionStart??input.value.length,end=input.selectionEnd??start,next=input.value.slice(0,start)+token+input.value.slice(end);target.set(next);requestAnimationFrame(()=>{input.focus();input.setSelectionRange(start+token.length,start+token.length);});};
 const preset=(kind:'center'|'right'|'total')=>setDraft(value=>({...value,footer:kind==='center'?{left:'',center:WORD_PAGE_TOKEN,right:''}:kind==='right'?{left:'',center:'',right:WORD_PAGE_TOKEN}:{left:'',center:`第 ${WORD_PAGE_TOKEN} 页，共 ${WORD_PAGES_TOKEN} 页`,right:''}}));
 const row=(scope:'main'|'first',part:Part,label:string)=><fieldset className="word-header-row" key={scope+part}><legend>{label}</legend>{SLOTS.map(([slot,name])=><label key={slot}>{name}<input aria-label={`${label}${name}`} autoFocus={scope==='main'&&part===focus&&slot==='center'} maxLength={255} value={line(scope,part)[slot]} disabled={readOnly} onFocus={event=>{last.current={input:event.currentTarget,set:text=>update(scope,part,slot,text)};}} onChange={event=>update(scope,part,slot,event.target.value)}/></label>)}</fieldset>;
 const save=()=>{if(readOnly)return;const empty=emptyWordLine(draft.header)&&emptyWordLine(draft.footer)&&(!draft.first||emptyWordLine(draft.first.header)&&emptyWordLine(draft.first.footer));
  try{setWordHeaderFooter(doc,empty&&!draft.first?undefined:draft,base);onClose();}catch(reason){setError(reason instanceof Error?reason.message:'页眉页脚保存失败');}};
 return <Modal title="页眉和页脚" onClose={onClose}><form className="word-header-dialog" onSubmit={event=>{event.preventDefault();save();}}>
  <p>页眉和页脚分左、中、右三处。插入的页码在打印和导出 DOCX 时按每页显示；编辑区只在首页上方与末页下方显示。</p>
  {row('main','header','页眉')}{row('main','footer','页脚')}
  <div className="word-header-tools" role="group" aria-label="插入域"><button type="button" disabled={readOnly} onMouseDown={event=>event.preventDefault()} onClick={()=>insert(WORD_PAGE_TOKEN)}>插入页码</button><button type="button" disabled={readOnly} onMouseDown={event=>event.preventDefault()} onClick={()=>insert(WORD_PAGES_TOKEN)}>插入总页数</button>
   <span>页码样式</span><button type="button" disabled={readOnly} onClick={()=>preset('center')}>页脚居中</button><button type="button" disabled={readOnly} onClick={()=>preset('right')}>页脚靠右</button><button type="button" disabled={readOnly} onClick={()=>preset('total')}>第 X 页，共 Y 页</button></div>
  <label><input type="checkbox" checked={!!draft.first} disabled={readOnly} onChange={event=>setDraft(value=>{const next={...value};if(event.target.checked)next.first={header:{...EMPTY_WORD_LINE},footer:{...EMPTY_WORD_LINE}};else delete next.first;return next;})}/>首页不同</label>
  {draft.first&&<>{row('first','header','首页页眉')}{row('first','footer','首页页脚')}</>}
  {readOnly&&<p role="status">当前文档只读。</p>}{error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button type="button" onClick={reload}>重新载入</button><button type="button" disabled={readOnly||!base} onClick={()=>{try{setWordHeaderFooter(doc,undefined,base);onClose();}catch(reason){setError(reason instanceof Error?reason.message:'删除失败');}}}>删除页眉页脚</button><button type="button" onClick={onClose}>取消</button><button className="primary" disabled={readOnly}>保存页眉页脚</button></div>
 </form></Modal>;
}
