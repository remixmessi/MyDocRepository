import {useState} from 'react';
import {defaultEditorSettings,type EditorSettings} from '@mydoc/editor-blocks';
import {Modal} from './dialogs';
export function EditorSettingsDialog({initial,onSave,onClose}:{initial:EditorSettings;onSave:(value:EditorSettings)=>Promise<void>;onClose:()=>void}){
 const [value,setValue]=useState(initial),[busy,setBusy]=useState(false),[error,setError]=useState('');
 return <Modal title="编辑器设置" onClose={()=>{if(!busy)onClose();}}><form onSubmit={event=>{event.preventDefault();setBusy(true);setError('');void onSave(value).then(onClose).catch(e=>setError(e instanceof Error?e.message:'设置保存失败，请重试')).finally(()=>setBusy(false));}}>
 <p className="organization-help">应用于当前工作区的所有正文，刷新后保留。</p><div className="editor-settings-fields">
 <label>正文宽度<select aria-label="正文宽度" value={value.width} disabled={busy} onChange={event=>setValue({...value,width:event.target.value as EditorSettings['width']})}><option value="normal">标准</option><option value="wide">宽</option><option value="full">全宽</option></select></label>
 <label>正文字号<select aria-label="正文字号" value={value.fontSize} disabled={busy} onChange={event=>setValue({...value,fontSize:Number(event.target.value) as EditorSettings['fontSize']})}>{[14,16,18,20].map(size=><option value={size} key={size}>{size} px</option>)}</select></label>
 <label>正文行高<select aria-label="正文行高" value={value.lineHeight} disabled={busy} onChange={event=>setValue({...value,lineHeight:Number(event.target.value) as EditorSettings['lineHeight']})}>{[1.5,1.8,2,2.2].map(height=><option value={height} key={height}>{height}</option>)}</select></label>
 <label className="settings-checkbox"><input type="checkbox" aria-label="显示正文选区工具栏" checked={value.selectionToolbar} disabled={busy} onChange={event=>setValue({...value,selectionToolbar:event.target.checked})}/>显示正文选区工具栏</label>
 <label className="settings-checkbox"><input type="checkbox" aria-label="显示移动端编辑工具栏" checked={value.mobileToolbar} disabled={busy} onChange={event=>setValue({...value,mobileToolbar:event.target.checked})}/>显示移动端编辑工具栏</label></div>
 <div className="editor-settings-preview" style={{fontSize:value.fontSize,lineHeight:value.lineHeight}}>写下今天的想法。<br/>预览正文的字号与行距。</div>{error&&<p role="alert" className="form-error">{error}</p>}<div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>setValue({...defaultEditorSettings})}>恢复默认</button><button type="button" disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy}>{busy?'正在保存…':'保存设置'}</button></div>
 </form></Modal>;
}
