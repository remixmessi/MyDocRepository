import {useRef,useState} from 'react';
import {createPortal} from 'react-dom';
import {Modal} from './dialogs';
import {cloudRequest} from '../../../common/nbstore/src/cloud';
import {isCloudWorkspaceArchive,type CloudWorkspaceArchive} from '../../../common/nbstore/src/workspace-archive';
import {downloadOfficeBytes} from './office-download';
import type {LocalWorkspace} from './workspace';

export async function downloadTeamWorkspaceBackup(workspace:LocalWorkspace,name:string){
 document.activeElement instanceof HTMLElement&&document.activeElement.blur();window.dispatchEvent(new Event('mydoc-before-close'));
 await workspace.prepareWorkspaceBackup();
 const response=await cloudRequest('/workspace-backup?'+new URLSearchParams({workspace:workspace.id}));
 downloadOfficeBytes(name+'.mydoc-team.json',new Uint8Array(await response.arrayBuffer()),'application/json');
}
export function TeamWorkspaceRestoreDialog({file,archive,onClose,onRestore}:{file:File;archive:CloudWorkspaceArchive;onClose:()=>void;onRestore:(id:string)=>Promise<void>}){
 const operation=useRef(crypto.randomUUID()),result=useRef<string>();const [busy,setBusy]=useState(false),[error,setError]=useState('');
 const valid=isCloudWorkspaceArchive(archive)&&typeof archive.workspace?.name==='string';
 return createPortal(<Modal title="恢复团队工作区" onClose={()=>{if(!busy)onClose();}}><form onSubmit={event=>{
  event.preventDefault();if(busy||!valid)return;setBusy(true);setError('');
  void(async()=>{
   if(!result.current){const response=await cloudRequest('/workspace-restore?'+new URLSearchParams({operationId:operation.current}),{method:'POST',headers:{'Content-Type':'application/json'},body:file});result.current=(await response.json() as {id:string}).id;}
   await onRestore(result.current!);
  })().catch(error=>setError(error instanceof Error?error.message:'工作区恢复失败')).finally(()=>setBusy(false));
 }}>
  <p className="office-backup-filename">{file.name}</p>
  {valid?<><p>备份工作区：{archive.workspace.name}</p><p>{archive.documents.length} 份协作状态 · {archive.office.length} 个办公文件 · {archive.assets.length} 个附件 · {archive.versions?.length??0} 个历史版本</p></>:<p role="alert" className="form-error">备份信息无效，请重新选择文件。</p>}
  <p>恢复到当前服务器上的新团队工作区，当前登录账户成为所有者。目录、回收状态、正文、办公内容及原件一并恢复；成员、授权、会话和个人设置不复制。</p>
  <p>备份最多 50 MB，服务器会校验内容。失败后可在此重试；恢复不会覆盖已有工作区。</p>
  {error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button type="button" disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||!valid}>{busy?'正在恢复…':error?'重试恢复':'恢复为新团队工作区'}</button></div>
 </form></Modal>,document.body);
}
