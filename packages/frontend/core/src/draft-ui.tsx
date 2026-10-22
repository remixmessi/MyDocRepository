import {storageRequest} from '../../../common/nbstore/src/sqlite';
import {useEffect,useRef,useState} from 'react';
import {Modal} from './dialogs';
import {FolderSelect} from './sidebar';
import {parseDraft,restoreDraft,MAX_DRAFT_SIZE,type DraftPreview,type DraftRestoreResult} from './draft';
import type {LocalWorkspace} from './workspace';
import {useFileDrop,FileDropHint} from '../../../editor/blocks/src/file-drop';

/**
 * 草稿恢复对话框：选择草稿文件、预览内容、恢复为当前个人工作区中的新文档。
 * 目标工作区必须是个人工作区；团队工作空间下由调用方先切换。
 */
export function DraftRestoreDialog({workspace,initialFile,onClose,onFinish,onBusyChange}:{workspace:LocalWorkspace;initialFile?:File;onClose:()=>void;onFinish:(id?:string)=>void;onBusyChange?:(busy:boolean)=>void}){
 const [preview,setPreview]=useState<DraftPreview>(),[selected,setSelected]=useState(new Set<string>()),[error,setError]=useState(''),[busy,setBusy]=useState(false),[results,setResults]=useState<DraftRestoreResult[]>(),[folderId,setFolderId]=useState('');
 const picker=useRef<HTMLInputElement>(null),alive=useRef(true),finished=useRef(false);
 useEffect(()=>{alive.current=true;return()=>{alive.current=false;};},[]);
 useEffect(()=>{onBusyChange?.(busy);return()=>onBusyChange?.(false);},[busy,onBusyChange]);
 const read=(file:File)=>{setError('');setResults(undefined);setPreview(undefined);
  if(file.size>MAX_DRAFT_SIZE){setError('草稿文件超过 50 MB，无法恢复');return;}
  void file.text().then(text=>{if(!alive.current)return;const value=parseDraft(text);setPreview(value);setSelected(new Set(value.documents.map(entry=>entry.sourceId)));}).catch(e=>{if(alive.current)setError(e instanceof Error?e.message:'无法读取草稿文件');});
 };
 useEffect(()=>{if(initialFile)read(initialFile);},[initialFile]);
 const {dragging,props:dropProps}=useFileDrop({disabled:busy||!!results,onFiles:files=>read(files[0]),onError:setError});
 const readStored=async()=>{setBusy(true);setError('');try{const rows=await (await storageRequest('/sync-drafts')).json() as {id:string;data:string}[];if(!rows.length)throw new Error('当前账户没有暂存文档草稿');const value=parseDraft(JSON.stringify({format:'mydoc-workspace-draft',version:1,workspaceId:'sqlite-drafts',documents:rows.map(row=>({id:row.id,update:Array.from(atob(row.data),c=>c.charCodeAt(0))}))}));if(alive.current){setResults(undefined);setPreview(value);setSelected(new Set(value.documents.map(entry=>entry.sourceId)));}}catch(e){if(alive.current)setError(e instanceof Error?e.message:'无法读取暂存草稿');}finally{if(alive.current)setBusy(false);}};
 const start=async()=>{if(!preview||busy)return;setBusy(true);setError('');
  try{const output=await restoreDraft(workspace,preview.documents.filter(entry=>selected.has(entry.sourceId)),folderId||null);if(alive.current)setResults(output);}
  catch(e){if(alive.current)setError(e instanceof Error?e.message:'草稿恢复失败，请重试');}
  finally{if(alive.current)setBusy(false);}
 };
 return <Modal title="从草稿恢复" wide onClose={()=>{if(!busy)onClose();}} restoreFocus={()=>!finished.current} drop={dropProps}>
  <p className="organization-help">读取状态区导出的工作空间草稿或单文档草稿。恢复会在当前个人工作区创建新文档，不写回原文档，也不恢复原工作空间的访问权限。草稿不含附件字节，缺失的文件会在恢复后提示。</p>
  {workspace.sync&&<p className="form-error" role="alert">当前是团队工作空间，草稿只能恢复到个人工作区。请先切换到个人工作区再恢复。</p>}
  {!results&&<FileDropHint dragging={dragging} disabled={busy} hint="支持导出的工作空间草稿或单文档草稿 JSON"/>}
  {!results&&<div className="import-file-actions"><button disabled={busy} onClick={()=>picker.current?.click()}>选择草稿文件</button><span>{preview?`已读取 ${preview.documents.length} 篇文档`:'尚未选择草稿'}</span><input ref={picker} hidden type="file" accept=".json,application/json" aria-label="选择草稿文件" onChange={event=>{const file=event.target.files?.[0];event.target.value='';if(file)read(file);}}/></div>}
  {error&&<p className="form-error" role="alert">{error}</p>}
  {!results&&<div className="import-file-actions"><button disabled={busy||!!workspace.sync} onClick={()=>void readStored()}>读取暂存草稿</button></div>}
  {preview&&!results&&<>
   <div className="import-options"><FolderSelect workspace={workspace} label="恢复目标文件夹" value={folderId} disabled={busy} onChange={setFolderId}/></div>
   <p role="status">{preview.kind==='workspace'?`工作空间草稿${preview.createdAt?` · 导出于 ${new Date(preview.createdAt).toLocaleString()}`:''}`:'单文档草稿'}</p>
   <div className="import-selection"><strong>{preview.documents.length} 篇文档 · 已选择 {selected.size} 篇</strong><button disabled={busy} onClick={()=>setSelected(new Set(preview.documents.map(entry=>entry.sourceId)))}>全选</button><button disabled={busy} onClick={()=>setSelected(new Set())}>全不选</button></div>
   <ul className="import-preview" aria-label="草稿恢复预览">{preview.documents.map(entry=><li key={entry.sourceId}><label><input type="checkbox" aria-label={`恢复 ${entry.title}`} checked={selected.has(entry.sourceId)} disabled={busy} onChange={event=>setSelected(current=>{const next=new Set(current);event.target.checked?next.add(entry.sourceId):next.delete(entry.sourceId);return next;})}/><span><strong>{entry.title}</strong><small>{entry.summary||'（正文为空）'}</small>{entry.surfaceElements>0&&<small>白板内容 {entry.surfaceElements} 个元素</small>}{entry.missingAssets.length>0&&<small className="form-error">草稿不含 {entry.missingAssets.length} 个附件的文件数据，恢复后需要重新上传</small>}</span></label></li>)}</ul>
  </>}
  {busy&&<p role="status">正在恢复草稿…</p>}
  {results&&<><p role="status">已恢复 {results.length} 篇文档到当前工作区。</p><ul className="import-preview" aria-label="草稿恢复结果">{results.map(result=><li key={result.id}><strong>{result.title}</strong>{result.missingAssets.length>0&&<small className="form-error">缺少 {result.missingAssets.length} 个附件文件，请重新上传</small>}</li>)}</ul></>}
  <div className="dialog-actions">{results?<button className="primary" onClick={()=>{finished.current=true;onFinish(results[results.length-1]?.id);}}>完成</button>:<><button disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||!preview||!selected.size||!!workspace.sync} onClick={()=>void start()}>{busy?'正在恢复…':'恢复为新文档'}</button></>}</div>
 </Modal>;
}
