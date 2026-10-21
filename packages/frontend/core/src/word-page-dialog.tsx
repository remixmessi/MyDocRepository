import {useState} from 'react';
import type {Y} from '@mydoc/crdt';
import type {WordPageLayout} from '../../../common/nbstore/src/office-word-page';
import {getWordPage,setWordPage} from './office-word-page';
import {Modal} from './dialogs';
const papers={a4:[11906,16838],letter:[12240,15840]} as const;
function paperName(page:WordPageLayout){return Object.entries(papers).find(([,size])=>Math.min(page.width,page.height)===size[0]&&Math.max(page.width,page.height)===size[1])?.[0]??'custom';}
export function WordPageDialog({doc,readOnly,onClose}:{doc:Y.Doc;readOnly:boolean;onClose:()=>void}){
 const [base,setBase]=useState(()=>getWordPage(doc)),[draft,setDraft]=useState(()=>getWordPage(doc)),[error,setError]=useState('');
 const margins=['top','right','bottom','left'] as const;
 const [values,setValues]=useState(()=>Object.fromEntries(margins.map(key=>[key,String(base[key]/20)])) as Record<typeof margins[number],string>);
 const reload=()=>{const page=getWordPage(doc);setBase(page);setDraft(page);setValues(Object.fromEntries(margins.map(key=>[key,String(page[key]/20)])) as typeof values);setError('');};
 return <Modal title="页面设置" onClose={onClose}><form onSubmit={event=>{
  event.preventDefault();if(readOnly)return;
  try{const page={...draft};for(const key of margins)page[key]=Math.round(Number(values[key])*20);setWordPage(doc,page,base);onClose();}catch(error){setError(error instanceof Error?error.message:'页面设置保存失败');}
 }}>
  <p>整篇共用一套页面设置。正文使用流式预览，导出保留纸张与页边距。</p>
  <label className="form-field">纸张<select aria-label="纸张" value={paperName(draft)} disabled={readOnly} onChange={event=>{const size=papers[event.target.value as keyof typeof papers];if(size)setDraft({...draft,width:draft.width>draft.height?size[1]:size[0],height:draft.width>draft.height?size[0]:size[1]});}}><option value="a4">A4</option><option value="letter">Letter</option>{paperName(draft)==='custom'&&<option value="custom">自定义（导入尺寸）</option>}</select></label>
  <label className="form-field">方向<select aria-label="方向" value={draft.width>draft.height?'landscape':'portrait'} disabled={readOnly} onChange={event=>{const short=Math.min(draft.width,draft.height),long=Math.max(draft.width,draft.height);setDraft({...draft,width:event.target.value==='landscape'?long:short,height:event.target.value==='landscape'?short:long});}}><option value="portrait">纵向</option><option value="landscape">横向</option></select></label>
  {margins.map((key,i)=><label className="form-field" key={key}>{['上','右','下','左'][i]}页边距（磅）<input type="number" required min="0" max="360" step="0.05" value={values[key]} disabled={readOnly} onChange={event=>setValues({...values,[key]:event.target.value})}/></label>)}
  {readOnly&&<p role="status">当前文档只读。</p>}{error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button type="button" onClick={reload}>重新载入</button><button type="button" onClick={onClose}>取消</button><button className="primary" disabled={readOnly}>保存页面设置</button></div>
 </form></Modal>;
}
