import {useEffect,useRef,useState} from 'react';
import {Modal} from './dialogs';
import {cloudJson} from '../../../common/nbstore/src/cloud';
import type {LocalWorkspace} from './workspace';

interface PermissionEntry {user:string;name:string;email:string;role:'viewer'|'editor'|'denied'}
export function DocumentPermissionDialog({workspace,documentId,title,onClose}:{workspace:LocalWorkspace;documentId:string;title:string;onClose:()=>void}){
 const [entries,setEntries]=useState<PermissionEntry[]>([]),[members,setMembers]=useState<{id:string;name:string;email:string;role:string}[]>([]),[busy,setBusy]=useState(true),[loaded,setLoaded]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState('');
 const doc='doc:'+documentId,alive=useRef(true),path='/document-permissions?workspace='+encodeURIComponent(workspace.id)+'&doc='+encodeURIComponent(doc);
 const load=async()=>{setBusy(true);setError('');setLoaded(false);try{const [value,list]=await Promise.all([cloudJson<PermissionEntry[]>(path),cloudJson<{id:string;name:string;email:string;role:string}[]>('/members?workspace='+encodeURIComponent(workspace.id))]);if(alive.current){setEntries(value);setMembers(list);setLoaded(true);}}catch(e){if(alive.current)setError(e instanceof Error?e.message:'无法读取文档权限');}finally{if(alive.current)setBusy(false);}};
 useEffect(()=>{alive.current=true;void load();return()=>{alive.current=false;};},[]);
 const submit=async(event:React.FormEvent<HTMLFormElement>)=>{event.preventDefault();const form=event.currentTarget,data=new FormData(form),role=data.get('role');if(role==='denied'&&!window.confirm('确定禁止该成员访问此文档？该成员将无法读取、同步或订阅此文档。'))return;setBusy(true);setError('');try{await cloudJson(path,'PUT',{doc,user:data.get('user'),role});form.reset();setNotice('文档权限已更新');await load();}catch(e){setError(e instanceof Error?e.message:'更新文档权限失败');setBusy(false);}};
 const remove=async(user:string)=>{if(!window.confirm('恢复空间默认权限？普通成员将恢复编辑权限，空间查看者仍为只读。'))return;setBusy(true);setError('');try{await cloudJson(path,'DELETE',{doc,user});setNotice('已恢复空间默认权限');await load();}catch(e){setError(e instanceof Error?e.message:'恢复默认权限失败');setBusy(false);}};
 const available=members.filter(member=>member.role==='member');
 return <Modal title={`文档权限 · ${title}`} wide onClose={()=>{if(!busy)onClose();}}>
  <p className="organization-help">默认权限：成员可编辑，查看者只读。所有者与管理员可管理文档。</p>
  {notice&&<p className="form-notice" role="status">{notice}</p>}
  {error&&<><p className="form-error" role="alert">{error}</p><button disabled={busy} onClick={()=>void load()}>重试</button></>}
  {busy&&!entries.length?<p role="status">正在读取文档权限…</p>:loaded&&<ul className="cloud-members" aria-label="文档授权成员">{entries.map(entry=><li key={entry.user}><div><strong>{entry.name}</strong><small>{entry.email}</small></div><span>{entry.role==='editor'?'可编辑':entry.role==='denied'?'禁止访问':'只读'}</span><button disabled={busy} onClick={()=>void remove(entry.user)}>恢复默认<span className="sr-only"> {entry.name}</span></button></li>)}{!entries.length&&<li className="panel-empty">所有成员使用空间默认权限。</li>}</ul>}
  <form className="cloud-settings-form" onSubmit={submit}><label className="form-field">成员<select name="user" required disabled={busy||!loaded} defaultValue=""><option value="" disabled>选择成员</option>{available.map(member=><option key={member.id} value={member.id}>{member.name} · {member.email}</option>)}</select></label><label className="form-field">权限<select name="role" disabled={busy||!loaded} defaultValue="viewer"><option value="viewer">只读</option><option value="editor">可编辑</option><option value="denied">禁止访问</option></select></label><button className="primary" disabled={busy||!loaded||!available.length}>保存文档权限</button></form>
  <div className="dialog-actions"><button disabled={busy} onClick={onClose}>完成</button></div>
 </Modal>;
}
