import {useRef,useState} from 'react';
import type {Y} from '@mydoc/crdt';
import {ImagePlus,Trash2} from 'lucide-react';
import {Modal} from './dialogs';
import {getWordHeaderFooter,getWordHeaderImages,setWordHeaderFooter,type WordHeaderPicture} from './office-word-page';
import {OfficeImage,saveOfficeImage} from './office-image-ui';
import type {LocalWorkspace} from './workspace';
import {EMPTY_WORD_LINE,WORD_PAGE_TOKEN,WORD_PAGES_TOKEN,WORD_HEADER_LINES,WORD_HEADER_IMAGE_MAX,emptyWordLine,wordHeaderImageKey,type WordHeaderFooter,type WordHeaderLine,type WordHeaderPart,type WordHeaderScope,type WordHeaderSlot} from '../../../common/nbstore/src/office-word-header';

const SLOTS=[['left','左'],['center','中'],['right','右']] as const;
const blank=():WordHeaderFooter=>({header:{...EMPTY_WORD_LINE},footer:{...EMPTY_WORD_LINE}});
const pair=()=>({header:{...EMPTY_WORD_LINE},footer:{...EMPTY_WORD_LINE}});
// Header pictures are fitted into a small box by default; the export limit is WORD_HEADER_IMAGE_MAX.
const IMAGE_FIT={width:160,height:48};
/**
 * 插入 → 页眉和页脚: left/centre/right text per line (up to five lines each), a picture per slot, page-number fields,
 * and different first and even pages, like Word.
 */
export function WordHeaderDialog({doc,workspace,id,readOnly,focus='header',onClose}:{doc:Y.Doc;workspace:LocalWorkspace;id:string;readOnly:boolean;focus?:WordHeaderPart;onClose:()=>void}){
 const [base,setBase]=useState(()=>getWordHeaderFooter(doc)),[draft,setDraft]=useState<WordHeaderFooter>(()=>base??blank()),[error,setError]=useState('');
 const [baseImages,setBaseImages]=useState(()=>getWordHeaderImages(doc)),[images,setImages]=useState<Record<string,WordHeaderPicture>>(baseImages),[busy,setBusy]=useState(false);
 const last=useRef<{input:HTMLTextAreaElement;set:(text:string)=>void}>(),picker=useRef<HTMLInputElement>(null),target=useRef<string>();
 const reload=()=>{const value=getWordHeaderFooter(doc),pictures=getWordHeaderImages(doc,value);setBase(value);setDraft(value??blank());setBaseImages(pictures);setImages(pictures);setError('');};
 const line=(scope:WordHeaderScope,part:WordHeaderPart):WordHeaderLine=>scope==='main'?draft[part]:draft[scope]![part];
 const update=(scope:WordHeaderScope,part:WordHeaderPart,slot:WordHeaderSlot,text:string)=>setDraft(value=>scope==='main'?{...value,[part]:{...value[part],[slot]:text}}:{...value,[scope]:{...value[scope]!,[part]:{...value[scope]![part],[slot]:text}}});
 // Field buttons insert at the caret of the last focused box, as Word's 页码 button does.
 const insert=(token:string)=>{const current=last.current;if(!current)return;const {input}=current,start=input.selectionStart??input.value.length,end=input.selectionEnd??start,next=input.value.slice(0,start)+token+input.value.slice(end);current.set(next);requestAnimationFrame(()=>{input.focus();input.setSelectionRange(start+token.length,start+token.length);});};
 const preset=(kind:'center'|'right'|'total')=>setDraft(value=>({...value,footer:kind==='center'?{left:'',center:WORD_PAGE_TOKEN,right:''}:kind==='right'?{left:'',center:'',right:WORD_PAGE_TOKEN}:{left:'',center:`第 ${WORD_PAGE_TOKEN} 页，共 ${WORD_PAGES_TOKEN} 页`,right:''}}));
 const choose=(key:string)=>{target.current=key;picker.current?.click();};
 const upload=async(file:File)=>{const key=target.current;if(!key||readOnly)return;setBusy(true);setError('');
  try{const {asset,width,height}=await saveOfficeImage(workspace,id,doc,'word',file,IMAGE_FIT);setImages(value=>({...value,[key]:{assetId:asset.id,width:Math.max(1,Math.min(WORD_HEADER_IMAGE_MAX.width,Math.round(width))),height:Math.max(1,Math.min(WORD_HEADER_IMAGE_MAX.height,Math.round(height))),alt:file.name.slice(0,1000)}}));}
  catch(reason){setError(reason instanceof Error?reason.message:'图片插入失败');}finally{setBusy(false);if(picker.current)picker.current.value='';}};
 const remove=(key:string)=>setImages(value=>{const next={...value};delete next[key];return next;});
 const row=(scope:WordHeaderScope,part:WordHeaderPart,label:string)=><fieldset className="word-header-row" key={scope+part}><legend>{label}</legend>{SLOTS.map(([slot,name])=>{
  const key=wordHeaderImageKey(scope,part,slot),picture=images[key];
  return <div className="word-header-slot" key={slot}><label>{name}<textarea aria-label={`${label}${name}`} autoFocus={scope==='main'&&part===focus&&slot==='center'} rows={2} maxLength={255} value={line(scope,part)[slot]} disabled={readOnly} onFocus={event=>{last.current={input:event.currentTarget,set:text=>update(scope,part,slot,text)};}} onChange={event=>update(scope,part,slot,event.target.value)}/></label>
   {picture?<div className="word-header-picture"><span className="word-header-thumb" style={{aspectRatio:`${picture.width} / ${picture.height}`}}><OfficeImage workspace={workspace} id={id} assetId={picture.assetId} alt={picture.alt}/></span><span>{picture.width} × {picture.height}</span><button type="button" aria-label={`删除图片（${label}·${name}）`} title="删除图片" disabled={readOnly||busy} onClick={()=>remove(key)}><Trash2 size={13} aria-hidden="true"/></button></div>
   :<button type="button" className="word-header-add-picture" aria-label={`插入图片（${label}·${name}）`} disabled={readOnly||busy} onClick={()=>choose(key)}><ImagePlus size={13} aria-hidden="true"/>图片</button>}
  </div>;})}</fieldset>;
 const toggle=(scope:'first'|'even',on:boolean)=>setDraft(value=>{const next={...value};if(on)next[scope]=pair();else delete next[scope];return next;});
 const save=()=>{if(readOnly||busy)return;const variants=[draft.first,draft.even].filter(Boolean) as {header:WordHeaderLine;footer:WordHeaderLine}[];
  const empty=emptyWordLine(draft.header)&&emptyWordLine(draft.footer)&&variants.every(value=>emptyWordLine(value.header)&&emptyWordLine(value.footer))&&!Object.keys(images).length;
  const tooLong=[draft.header,draft.footer,...variants.flatMap(value=>[value.header,value.footer])].some(value=>Object.values(value).some(text=>text.split('\n').length>WORD_HEADER_LINES));
  if(tooLong){setError(`每处最多 ${WORD_HEADER_LINES} 行`);return;}
  try{setWordHeaderFooter(doc,empty&&!variants.length?undefined:draft,base,images,baseImages);onClose();}catch(reason){setError(reason instanceof Error?reason.message:'页眉页脚保存失败');}};
 return <Modal title="页眉和页脚" onClose={onClose}><form className="word-header-dialog" onSubmit={event=>{event.preventDefault();save();}}>
  <p>页眉和页脚分左、中、右三处，每处最多 {WORD_HEADER_LINES} 行（按 Enter 换行），可各放一张图片，显示在该处文字上方。插入的页码在打印和导出 DOCX 时按每页显示；编辑区只在首页上方与末页下方显示。</p>
  <input ref={picker} type="file" accept="image/png,image/jpeg" aria-label="选择页眉页脚图片" hidden onChange={event=>{const file=event.target.files?.[0];if(file)void upload(file);}}/>
  {row('main','header',draft.even?'奇数页页眉':'页眉')}{row('main','footer',draft.even?'奇数页页脚':'页脚')}
  <div className="word-header-tools" role="group" aria-label="插入域"><button type="button" disabled={readOnly} onMouseDown={event=>event.preventDefault()} onClick={()=>insert(WORD_PAGE_TOKEN)}>插入页码</button><button type="button" disabled={readOnly} onMouseDown={event=>event.preventDefault()} onClick={()=>insert(WORD_PAGES_TOKEN)}>插入总页数</button>
   <span>页码样式</span><button type="button" disabled={readOnly} onClick={()=>preset('center')}>页脚居中</button><button type="button" disabled={readOnly} onClick={()=>preset('right')}>页脚靠右</button><button type="button" disabled={readOnly} onClick={()=>preset('total')}>第 X 页，共 Y 页</button></div>
  <div className="word-header-options"><label><input type="checkbox" checked={!!draft.first} disabled={readOnly} onChange={event=>toggle('first',event.target.checked)}/>首页不同</label><label><input type="checkbox" checked={!!draft.even} disabled={readOnly} onChange={event=>toggle('even',event.target.checked)}/>奇偶页不同</label></div>
  {draft.first&&<>{row('first','header','首页页眉')}{row('first','footer','首页页脚')}</>}
  {draft.even&&<>{row('even','header','偶数页页眉')}{row('even','footer','偶数页页脚')}</>}
  {busy&&<p role="status">正在插入图片…</p>}{readOnly&&<p role="status">当前文档只读。</p>}{error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button type="button" onClick={reload}>重新载入</button><button type="button" disabled={readOnly||busy||!base} onClick={()=>{try{setWordHeaderFooter(doc,undefined,base,{},baseImages);onClose();}catch(reason){setError(reason instanceof Error?reason.message:'删除失败');}}}>删除页眉页脚</button><button type="button" onClick={onClose}>取消</button><button className="primary" disabled={readOnly||busy}>保存页眉页脚</button></div>
 </form></Modal>;
}
