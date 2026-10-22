import {createPortal} from 'react-dom';
import {useRef,useState} from 'react';
import type {LocalWorkspace} from './workspace';
import {Modal} from './dialogs';
import {cloudRequest} from '../../../common/nbstore/src/cloud';
import {downloadOfficeBytes} from './office-download';
import {useFileDrop,FileDropHint} from '../../../editor/blocks/src/file-drop';

export function OfficeBackupDialog({workspace,id,initialFile,onClose,onFinish}:{workspace:LocalWorkspace;id?:string;initialFile?:File;onClose:()=>void;onFinish:(id:string)=>void}){
 const operation=useRef(crypto.randomUUID());
 const [file,setFile]=useState<File|undefined>(initialFile),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const choose=(selected?:File)=>{operation.current=crypto.randomUUID();setFile(selected);setError(selected&&!/\.json$/i.test(selected.name)?'请选择 .mydoc-office.json 办公文件备份':'');};
 const {dragging,props:dropProps}=useFileDrop({disabled:!!id||busy||workspace.readOnly,onFiles:files=>choose(files[0]),onError:setError});
 return createPortal(<Modal title={id?'备份办公文件':'从备份恢复办公文件'} onClose={()=>{if(!busy)onClose();}} drop={id?undefined:dropProps}><form onSubmit={event=>{
  event.preventDefault();if(busy)return;setBusy(true);setError('');
  void(async()=>{
   if(id){await workspace.prepareOfficeBackup(id);const response=await cloudRequest('/office-backup?'+new URLSearchParams({workspace:workspace.id,id}));downloadOfficeBytes(workspace.resourceTitle(id)+'.mydoc-office.json',new Uint8Array(await response.arrayBuffer()),'application/json');onClose();}
   else{if(!file)throw new Error('请选择办公文件备份');if(file.size>50*1024*1024)throw new Error('备份超过 50 MB');if(workspace.readOnly)throw new Error('当前工作区只读');onFinish(await workspace.restoreOfficeBackup(file,operation.current));}
  })().catch(error=>setError(error instanceof Error?error.message:'办公文件备份失败')).finally(()=>setBusy(false));
 }}>
  <p>{id?'备份包含当前已确认的协作内容、引用的办公图片及原件，不包含成员授权、在线选区或版本历史。未提交的表格草稿需先处理。':'恢复为当前工作区中的新文件，不覆盖已有文件。新文件沿用目标工作区的默认成员权限；备份校验值用于检测损坏，不证明来源可信。'}</p>
  <p>这是单文件备份，不包含整个工作区、笔记目录或跨资源媒体。每份备份最多 50 MB。</p>
  {!id&&<FileDropHint dragging={dragging} disabled={busy||workspace.readOnly} hint="支持 .mydoc-office.json 单文件备份"/>}
  {!id&&<label className="form-field">办公备份文件<input type="file" accept=".json" disabled={busy||workspace.readOnly} onChange={event=>choose(event.target.files?.[0])}/></label>}
  {file&&<p className="office-backup-filename">{file.name}</p>}{error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button type="button" disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||!id&&(!file||workspace.readOnly)}>{busy?'正在处理…':id?'下载备份':'恢复为新文件'}</button></div>
 </form></Modal>,document.body);
}
