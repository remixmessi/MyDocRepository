import {useEffect,useId,useState} from 'react';
import {ArrowLeft,ChevronLeft,ChevronRight,RefreshCw} from 'lucide-react';
import {cloudJson,type CloudWorkspace} from '../../../common/nbstore/src/cloud';

interface ServerUser {id:string;email:string;name:string;verified:number;createdAt:number;serverAdmin:number;defaultAdmin:number}
interface UserPage {users:ServerUser[];total:number;offset:number;limit:number}
const formatDate=(value:number)=>value?new Date(value).toLocaleString('zh-CN',{hour12:false}):'—';

export function ServerAdminSettings({accountId,busy,run,onBack,workspaces,activeWorkspaceId}:{accountId:string;busy:boolean;run:(work:()=>Promise<void>)=>void;onBack:()=>void;workspaces:CloudWorkspace[];activeWorkspaceId?:string}){
 const [workspace,setWorkspace]=useState(()=>workspaces.some(space=>space.id===activeWorkspaceId)?activeWorkspaceId!:'');
 const workspaceSelectId=useId();
 const [page,setPage]=useState<UserPage>(),[offset,setOffset]=useState(0),[query,setQuery]=useState(''),[search,setSearch]=useState(''),[error,setError]=useState(''),[notice,setNotice]=useState(''),[loading,setLoading]=useState(true),[refresh,setRefresh]=useState(0),[confirm,setConfirm]=useState<ServerUser>();
 useEffect(()=>{if(!workspace){setPage(undefined);setLoading(false);return;}let alive=true;setPage(undefined);setLoading(true);setError('');void cloudJson<UserPage>('/admin/users?workspace='+encodeURIComponent(workspace)+'&offset='+offset+'&q='+encodeURIComponent(search)).then(value=>{if(alive)setPage(value);}).catch(error=>{if(alive)setError(error instanceof Error?error.message:'无法读取用户列表');}).finally(()=>{if(alive)setLoading(false);});return()=>{alive=false;};},[workspace,offset,search,refresh]);
 const update=(user:ServerUser)=>run(async()=>{await cloudJson('/admin/user','PATCH',{user:user.id,serverAdmin:!user.serverAdmin});setNotice(user.serverAdmin?'已取消服务器管理员权限':'已授予服务器管理员权限');setConfirm(undefined);setRefresh(value=>value+1);});
 return <section className="account-settings server-admin-settings">
  <header className="cloud-settings-head"><button className="workspace-back" disabled={busy} onClick={onBack}><ArrowLeft size={16}/>返回账户设置</button><h3>服务器用户管理</h3><button title="刷新用户列表" aria-label="刷新用户列表" disabled={busy||loading} onClick={()=>setRefresh(value=>value+1)}><RefreshCw size={16}/></button></header>
  <div className="form-field"><label htmlFor={workspaceSelectId}>管理工作空间</label><select id={workspaceSelectId} value={workspace} disabled={busy} onChange={event=>{setWorkspace(event.target.value);setOffset(0);setConfirm(undefined);setNotice('');setError('');}}><option value="">请选择团队工作空间</option>{workspaces.map(space=><option key={space.id} value={space.id}>{space.name}</option>)}</select></div>
  <p className="organization-help">{workspace?'显示所选空间的用户。服务器管理员权限适用于整个服务器，空间成员角色请在团队工作空间中管理。':workspaces.length?'选择工作空间后查看和管理用户。':'当前没有已加入的团队工作空间，请先创建或加入空间。'}</p>
  {notice&&<p role="status" className="form-notice">{notice}</p>}{error&&<p role="alert" className="form-error">{error}<button disabled={busy} onClick={()=>setRefresh(value=>value+1)}>重试</button></p>}
  {workspace&&<form onSubmit={event=>{event.preventDefault();setOffset(0);setSearch(query.trim());}}><label className="form-field">搜索用户<input value={query} onChange={event=>setQuery(event.target.value)} maxLength={254} disabled={busy}/></label><button disabled={busy||loading}>搜索</button></form>}
  {workspace&&(loading?<p role="status">正在读取用户列表…</p>:!error&&<>
   <ul className="cloud-members" aria-label="服务器用户列表">{page?.users.map(user=><li key={user.id}>
    <div><strong>{user.name}{user.id===accountId?'（当前账户）':''}</strong><small>{user.email}</small><small>注册于 {formatDate(user.createdAt)} · {user.verified?'邮箱已验证':'邮箱未验证'}</small></div>
    <span>{user.defaultAdmin?'默认管理员':user.serverAdmin?'服务器管理员':'普通用户'}</span>
    <button disabled={busy||user.id===accountId||!!user.defaultAdmin} onClick={()=>setConfirm(user)}>{user.serverAdmin?'取消服务器管理员':'设为服务器管理员'}<span className="sr-only"> {user.name}</span></button>
   </li>)}{page?.users.length===0&&<li className="panel-empty">没有符合条件的用户。</li>}</ul>
   {page&&<div className="dialog-actions"><span role="status">共 {page.total} 个用户 · 第 {Math.floor(offset/page.limit)+1} 页</span><button title="上一页" aria-label="上一页" disabled={busy||offset===0} onClick={()=>setOffset(Math.max(0,offset-page.limit))}><ChevronLeft size={16}/></button><button title="下一页" aria-label="下一页" disabled={busy||offset+page.limit>=page.total} onClick={()=>setOffset(offset+page.limit)}><ChevronRight size={16}/></button></div>}
  </>)}
  {confirm&&<div className="cloud-confirm" role="alertdialog" aria-label="确认管理员权限变更"><p>{confirm.serverAdmin?'取消':'授予'}“{confirm.name}”的服务器管理员权限？此变更适用于整个服务器。</p><div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setConfirm(undefined)}>取消</button><button className="primary" disabled={busy} onClick={()=>update(confirm)}>确认变更</button></div></div>}
 </section>;
}
