import {useEffect,useLayoutEffect,useRef,useState,type RefObject} from 'react';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {ListOrdered,Trash2} from 'lucide-react';
import {Modal} from './dialogs';
import {wordTextIdentity} from './office-word-flow';
import {applyOfficeTextEdit} from './office-text';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import {wordVisibleText} from '../../../common/nbstore/src/office-word-revision';
import {wordNoteRefs,wordNotes,wordTocLevels,wordWatermarkSize,validWordWatermark,WORD_NOTE_CHAR,WORD_WATERMARK_PRESETS,DEFAULT_WORD_WATERMARK_COLOR,type WordWatermark,type WordTocLevels} from '../../../common/nbstore/src/office-word-reference';
import type {WordPageLayout} from '../../../common/nbstore/src/office-word-page';
import type {WordPageBreak} from './use-word-pagination';

/** Headings (levels 1–3) as they read: tracked deletions and note references left out. */
export function wordHeadingEntries(doc:Y.Doc,levels=3){
 return doc.getArray<WordBlock>('paragraphs').toArray().flatMap(block=>{if(!(block instanceof Y.Text))return [];const level=Number(block.getAttribute('heading')??0),title=wordVisibleText(block).replaceAll(WORD_NOTE_CHAR,'').replace(/\s+/g,' ').trim();return level>=1&&level<=levels&&title?[{text:block,level,title}]:[];});
}
const pageOf=(top:number,breaks:WordPageBreak[])=>breaks.filter(entry=>entry.top<=top+1).length+1;
/** 目录: live entries from the headings with page numbers from the editor's estimated pagination; links jump to the heading. */
export function WordTocBlock({block,doc,paper,breaks,readOnly,onLocate,onDelete}:{block:Y.Map<unknown>;doc:Y.Doc;paper:RefObject<HTMLElement>;breaks:WordPageBreak[];readOnly:boolean;onLocate:(text:Y.Text)=>void;onDelete:()=>void}){
 const levels=wordTocLevels(block),entries=wordHeadingEntries(doc,levels),[pages,setPages]=useState<number[]>([]),[focused,setFocused]=useState(false);
 // Measured after layout: the heading's position against the page breaks.
 useLayoutEffect(()=>{const root=paper.current;if(!root)return;const next=entries.map(entry=>{const element=root.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(entry.text)}"]`)?.closest<HTMLElement>('.office-word-paragraph');return element?pageOf(element.offsetTop,breaks):1;});setPages(current=>JSON.stringify(current)===JSON.stringify(next)?current:next);});
 return <section className="office-word-toc" aria-label="目录" tabIndex={0} onFocus={()=>setFocused(true)} onBlur={event=>{if(!event.currentTarget.contains(event.relatedTarget as Node))setFocused(false);}} onKeyDown={event=>{if(!readOnly&&event.target===event.currentTarget&&(event.key==='Delete'||event.key==='Backspace')){event.preventDefault();onDelete();}}}>
  {focused&&!readOnly&&<div className="word-toc-tools" role="toolbar" aria-label="目录工具"><label>级别<select aria-label="目录级别" value={levels} onChange={event=>doc.transact(()=>block.set('levels',Number(event.target.value) as WordTocLevels),LOCAL_ORIGIN)}>{[1,2,3].map(value=><option key={value} value={value}>{value} 级目录</option>)}</select></label><button type="button" onClick={onDelete}><Trash2 size={14} aria-hidden="true"/>删除目录</button></div>}
  <h3 className="word-toc-title">目录</h3>
  {entries.length?<ol>{entries.map((entry,i)=><li key={wordTextIdentity(entry.text)??i} data-level={entry.level}><a href="#" onClick={event=>{event.preventDefault();onLocate(entry.text);}}><span className="word-toc-text">{entry.title}</span><span className="word-toc-leader" aria-hidden="true"/><span className="word-toc-page" aria-label={'第 '+(pages[i]??1)+' 页'}>{pages[i]??1}</span></a></li>)}</ol>:<p className="word-toc-empty">未找到目录项。在“开始”中把段落设为标题 1–{levels} 后，目录会自动更新。</p>}
 </section>;
}
/** One note's text, edited as plain text and merged character by character with other members' edits. */
function WordNoteField({text,label,readOnly,assertWritable}:{text:Y.Text;label:string;readOnly:boolean;assertWritable:()=>void}){
 const [value,setValue]=useState(text.toString()),ref=useRef<HTMLTextAreaElement>(null),caret=useRef<[number,number]>();
 useEffect(()=>{const update=()=>{const element=ref.current;if(element&&document.activeElement===element)caret.current=[element.selectionStart,element.selectionEnd];setValue(text.toString());};text.observe(update);update();return()=>text.unobserve(update);},[text]);
 useLayoutEffect(()=>{const element=ref.current,saved=caret.current;if(element&&saved){caret.current=undefined;element.setSelectionRange(Math.min(saved[0],value.length),Math.min(saved[1],value.length));}},[value]);
 return <><textarea ref={ref} className="word-note-input" aria-label={label} rows={1} maxLength={5000} readOnly={readOnly} value={value} onChange={event=>{if(readOnly)return;try{assertWritable();applyOfficeTextEdit(text,text.toString(),event.target.value.replaceAll(WORD_NOTE_CHAR,'').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\u2029]/g,''));}catch{setValue(text.toString());}}}/><span className="word-note-print">{value}</span></>;
}
/** 脚注 and 尾注 listed after the body (the editor's draft view); DOCX export places them per page and at the end, as Word does. */
export function WordNotesSection({doc,readOnly,assertWritable,onLocate}:{doc:Y.Doc;readOnly:boolean;assertWritable:()=>void;onLocate:(text:Y.Text,offset:number)=>void}){
 const refs=wordNoteRefs(doc,wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray())),seen=new Set<string>(),notes=wordNotes(doc),unique=refs.filter(ref=>!seen.has(ref.id)&&seen.add(ref.id));
 if(!unique.length)return null;
 const list=(kind:'footnote'|'endnote')=>{const items=unique.filter(ref=>ref.kind===kind);if(!items.length)return null;return <section className="word-notes-group" aria-label={kind==='footnote'?'脚注':'尾注'}><h4>{kind==='footnote'?'脚注':'尾注'}</h4><ol>{items.map(ref=>{const text=notes?.get(ref.id)?.get('text');return text instanceof Y.Text?<li key={ref.id} data-word-note={ref.id} data-deleted={ref.deleted||undefined}><a href="#" className="word-note-label" title="转到引用位置" onClick={event=>{event.preventDefault();onLocate(ref.text,ref.offset);}}>{ref.label}</a><WordNoteField text={text} label={`${kind==='footnote'?'脚注':'尾注'} ${ref.label}`} readOnly={readOnly} assertWritable={assertWritable}/></li>:null;})}</ol></section>;};
 return <div className="word-notes" contentEditable={false}>{list('footnote')}{list('endnote')}</div>;
}
/** Number every reference in the body (arabic for footnotes, roman for endnotes) after the paragraphs re-render. */
export function useWordNoteLabels(root:RefObject<HTMLElement>,doc:Y.Doc){
 useEffect(()=>{
  const label=()=>{const element=root.current;if(!element)return;const refs=wordNoteRefs(doc,wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray())),labels=new Map(refs.map(ref=>[ref.id,ref.label]));for(const node of element.querySelectorAll<HTMLElement>('.word-note-ref')){const value=labels.get(node.dataset.note??'')??'';if(node.dataset.label!==value)node.dataset.label=value;}};
  // Paragraph observers re-render first, then this deep observer labels the new spans.
  // New reference spans (a paragraph re-rendered after typing or a remote edit) are labelled as they appear.
  const paragraphs=doc.getArray('paragraphs'),notes=doc.getMap('wordNotes'),mutation=new MutationObserver(records=>{if(records.some(record=>[...record.addedNodes].some(node=>node instanceof HTMLElement&&(node.matches('.word-note-ref')||!!node.querySelector('.word-note-ref')))))label();});
  if(root.current)mutation.observe(root.current,{childList:true,subtree:true});paragraphs.observeDeep(label);notes.observe(label);label();const frame=requestAnimationFrame(label);
  return()=>{mutation.disconnect();paragraphs.unobserveDeep(label);notes.unobserve(label);cancelAnimationFrame(frame);};
 },[root,doc]);
}
/** One watermark centred on each estimated page, behind the text. */
export function WordWatermarkLayer({watermark,page,breaks,height}:{watermark:WordWatermark;page:WordPageLayout;breaks:WordPageBreak[];height:number}){
 const tops=[0,...breaks.map(entry=>entry.top)],size=wordWatermarkSize(watermark,page);
 return <div className="word-watermark-layer" aria-hidden="true">{tops.map((top,i)=>{const bottom=i+1<tops.length?tops[i+1]:Math.max(height,top+page.height/15);return <div key={i} className="word-watermark" style={{top,height:bottom-top}}><span style={{fontSize:size+'pt',color:watermark.color,transform:watermark.layout==='diagonal'?'rotate(-45deg)':undefined}}>{watermark.text}</span></div>;})}</div>;
}
/** 水印 dialog: preset or custom text, colour, size and layout; semi-transparent like WPS's default. */
export function WordWatermarkDialog({current,readOnly,onSave,onClose}:{current?:WordWatermark;readOnly:boolean;onSave:(value:WordWatermark|undefined)=>void;onClose:()=>void}){
 const [text,setText]=useState(current?.text??WORD_WATERMARK_PRESETS[0]),[color,setColor]=useState(current?.color??DEFAULT_WORD_WATERMARK_COLOR),[layout,setLayout]=useState<WordWatermark['layout']>(current?.layout??'diagonal'),[size,setSize]=useState(current?.size?String(current.size):'auto'),[error,setError]=useState('');
 const value:WordWatermark={text:text.trim(),color:color.toUpperCase(),layout,...(size!=='auto'?{size:Number(size)}:{})};
 return <Modal title="水印" onClose={onClose}>
  <label className="form-field">文字<input aria-label="水印文字" list="word-watermark-presets" maxLength={50} value={text} onChange={event=>setText(event.target.value)}/></label><datalist id="word-watermark-presets">{WORD_WATERMARK_PRESETS.map(preset=><option key={preset} value={preset}/>)}</datalist>
  <div className="word-watermark-presets" role="group" aria-label="常用水印">{WORD_WATERMARK_PRESETS.slice(0,6).map(preset=><button key={preset} type="button" aria-pressed={text===preset} onClick={()=>setText(preset)}>{preset}</button>)}</div>
  <label className="form-field">颜色<input type="color" aria-label="水印颜色" value={color} onChange={event=>setColor(event.target.value.toUpperCase())}/></label>
  <label className="form-field">字号<select aria-label="水印字号" value={size} onChange={event=>setSize(event.target.value)}><option value="auto">自动</option>{[36,48,54,60,72,90,105,120,144].map(value=><option key={value} value={value}>{value}</option>)}</select></label>
  <fieldset><legend>版式</legend><label><input type="radio" name="word-watermark-layout" checked={layout==='diagonal'} onChange={()=>setLayout('diagonal')}/>斜式</label><label><input type="radio" name="word-watermark-layout" checked={layout==='horizontal'} onChange={()=>setLayout('horizontal')}/>水平</label></fieldset>
  <p className="dialog-note">水印半透明显示在每页正文之后，打印及导出 DOCX 时写入页眉。</p>
  {error&&<p role="alert">{error}</p>}
  <div className="dialog-actions">{current&&<button disabled={readOnly} onClick={()=>onSave(undefined)}>删除水印</button>}<button onClick={onClose}>取消</button><button className="primary" disabled={readOnly} onClick={()=>{if(!validWordWatermark(value)){setError('请填写 1–50 个字符的水印文字');return;}onSave(value);}}>确定</button></div>
 </Modal>;
}
