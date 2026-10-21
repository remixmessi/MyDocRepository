import {useState} from 'react';
import {Modal} from './dialogs';
import type {LocalWorkspace} from './workspace';

const TYPES=[
 {id:'note',label:'笔记',extension:''},
 {id:'docx',label:'Word 文档（.docx）',extension:'.docx'},
 {id:'xlsx',label:'Excel 表格（.xlsx）',extension:'.xlsx'},
 {id:'pptx',label:'PowerPoint 演示（.pptx）',extension:'.pptx'},
 {id:'pdf',label:'PDF（空白页面与批注）',extension:'.pdf'},
 {id:'csv',label:'CSV 表格（.csv）',extension:'.csv'},
 {id:'tsv',label:'TSV 表格（.tsv）',extension:'.tsv'},
 {id:'txt',label:'纯文本（.txt）',extension:'.txt'},
];
export function NewFileDialog({workspace,folderId,officeOnly=false,onClose,onCreated}:{workspace:LocalWorkspace;folderId?:string;officeOnly?:boolean;onClose:()=>void;onCreated:(id:string)=>void}){
 const [type,setType]=useState(officeOnly?'docx':'note'),[name,setName]=useState('未命名');
 const [busy,setBusy]=useState(false),[error,setError]=useState('');
 const selected=TYPES.find(item=>item.id===type)!;
 const readOnly=workspace.readOnly;
 return <Modal title="新建文件" onClose={()=>{if(!busy)onClose();}}><form onSubmit={event=>{
  event.preventDefault();if(busy||readOnly||!name.trim())return;
  setBusy(true);setError('');
  const base=name.trim();const filename=selected.extension&&!base.toLowerCase().endsWith(selected.extension)?base+selected.extension:base;
  void (type==='note'?workspace.create(name.trim(),undefined,folderId??null):workspace.createOfficeFile(filename,folderId))
   .then(onCreated).catch(e=>setError(e instanceof Error?e.message:'创建失败，请重试')).finally(()=>setBusy(false));
 }}>
  <label className="form-field">文件类型<select aria-label="文件类型" autoFocus value={type} disabled={busy||readOnly} onChange={event=>setType(event.target.value)}>
   {TYPES.filter(item=>!officeOnly||item.id!=='note').map(item=><option key={item.id} value={item.id} disabled={item.id!=='note'&&!workspace.officeAvailable}>{item.label}</option>)}
  </select></label>
  <label className="form-field">文件名称<input aria-label="文件名称" value={name} onChange={event=>setName(event.target.value)} maxLength={160} required disabled={busy||readOnly}/></label>
  {type!=='note'&&<p className="organization-help">文件将显示在当前文件夹中，与笔记统一管理{selected.extension?`，扩展名 ${selected.extension} 会自动补齐`:''}。{type==='pdf'?'创建一页空白 PDF，可添加批注。':''}</p>}
  {!workspace.officeAvailable&&<p className="organization-help">办公文件需要本地 SQLite 存储服务。</p>}
  {readOnly&&<p role="status">当前工作空间为只读，无法创建文件。</p>}
  {error&&<p className="form-error" role="alert">{error}</p>}
  <div className="dialog-actions"><button type="button" disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||readOnly||!name.trim()}>{busy?'正在创建…':'创建'}</button></div>
 </form></Modal>;
}
