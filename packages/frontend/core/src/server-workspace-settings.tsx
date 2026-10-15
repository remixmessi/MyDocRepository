import {useEffect,useState} from 'react';
import {cloudJson,cloudServerUrl} from '../../../common/nbstore/src/cloud';

type Role='owner'|'admin'|'member'|'viewer';
interface Member {id:string;name:string;email:string;role:Role}
interface Space {id:string;name:string;createdAt:number;memberCount:number;ownerName:string|null;ownerEmail:string|null}
interface SpacePage {spaces:Space[];total:number;limit:number}
interface Detail {id:string;name:string;createdAt:number;members:Member[];documents:number;storageBytes:number}
interface AuditPage {entries:{id:string;actorName:string;action:string;details:string;createdAt:number}[];total:number;limit:number}
type Run=(work:()=>Promise<void>)=>void;
const roles:Record<Role,string>={owner:'所有者',admin:'管理员',member:'成员',viewer:'查看者'};
const date=(value:number)=>new Date(value).toLocaleString('zh-CN',{hour12:false});
const path=(route:string,id:string)=>route+'?workspace='+encodeURIComponent(id);

export function ServerWorkspaceSettings({busy,run,onChanged}:{busy:boolean;run:Run;onChanged:()=>Promise<void>}){
 const [page,setPage]=useState<SpacePage>(),[offset,setOffset]=useState(0),[query,setQuery]=useState(''),[search,setSearch]=useState(''),[selected,setSelected]=useState<string>(),[refresh,setRefresh]=useState(0),[loading,setLoading]=useState(true),[error,setError]=useState('');
 useEffect(()=>{let alive=true;setLoading(true);setError('');setPage(undefined);void cloudJson<SpacePage>('/admin/workspaces?offset='+offset+'&q='+encodeURIComponent(search)).then(value=>{if(alive){if(offset>0&&value.spaces.length===0){setOffset(0);return;}setPage(value);}}).catch(e=>{if(alive)setError(e instanceof Error?e.message:'无法读取空间列表');}).finally(()=>{if(alive)setLoading(false);});return()=>{alive=false;};},[offset,search,refresh]);
 if(selected)return <WorkspaceAdminDetail key={selected} id={selected} busy={busy} run={run} onBack={()=>{setSelected(undefined);setRefresh(value=>value+1);}} onChanged={onChanged}/>;
 return <section className="server-workspace-settings"><h2>所有团队空间</h2><p className="organization-help">当前团队服务器：{cloudServerUrl()||location.origin}</p><p className="organization-help">管理当前服务器的全部团队空间。这里的管理权限不会自动授予文档和附件访问权限。</p>
  <form className="global-space-toolbar" onSubmit={event=>{event.preventDefault();setOffset(0);setSearch(query.trim());setRefresh(value=>value+1);}}><label className="form-field">搜索空间<input value={query} maxLength={160} disabled={busy} onChange={event=>setQuery(event.target.value)}/></label><button disabled={busy||loading}>搜索</button><button type="button" disabled={busy||loading} onClick={()=>setRefresh(value=>value+1)}>刷新</button></form>
  {error&&<p role="alert" className="form-error">{error}<button disabled={busy} onClick={()=>setRefresh(value=>value+1)}>重试</button></p>}
  {loading?<p role="status">正在读取所有团队空间…</p>:!error&&<><ul className="cloud-members global-space-list" aria-label="所有团队空间">{page?.spaces.map(space=><li key={space.id}><div><strong>{space.name}</strong><small>所有者：{space.ownerName??'未指定'}{space.ownerEmail?' · '+space.ownerEmail:''}</small><small>{space.memberCount} 位成员 · 创建于 {date(space.createdAt)}</small></div><button disabled={busy} onClick={()=>setSelected(space.id)}>管理<span className="sr-only"> {space.name}</span></button></li>)}</ul>{page?.spaces.length===0&&<p role="status">没有符合条件的团队空间。可调整搜索条件或刷新列表。</p>}{page&&<Pagination offset={offset} total={page.total} limit={page.limit} busy={busy} change={setOffset}/>}</>}
  <AuditHistory busy={busy} refresh={refresh}/>
 </section>;
}
function Pagination({offset,total,limit,busy,change}:{offset:number;total:number;limit:number;busy:boolean;change:(offset:number)=>void}){return <div className="dialog-actions global-space-pagination"><span>共 {total} 条 · 第 {Math.floor(offset/limit)+1} 页</span><button disabled={busy||offset===0} onClick={()=>change(Math.max(0,offset-limit))}>上一页</button><button disabled={busy||offset+limit>=total} onClick={()=>change(offset+limit)}>下一页</button></div>;}

type Pending={kind:'delete'}|{kind:'owner';user:Member}|{kind:'remove';user:Member}|{kind:'member';user:Member;role:Role;add:boolean};
function WorkspaceAdminDetail({id,busy,run,onBack,onChanged}:{id:string;busy:boolean;run:Run;onBack:()=>void;onChanged:()=>Promise<void>}){
 const [detail,setDetail]=useState<Detail>(),[refresh,setRefresh]=useState(0),[name,setName]=useState(''),[error,setError]=useState(''),[notice,setNotice]=useState(''),[loading,setLoading]=useState(true),[pending,setPending]=useState<Pending>(),[confirmName,setConfirmName]=useState('');
 useEffect(()=>{let alive=true;setLoading(true);setError('');setDetail(undefined);void cloudJson<Detail>(path('/admin/workspace',id)).then(value=>{if(alive){setDetail(value);setName(value.name);}}).catch(e=>{if(alive)setError(e instanceof Error?e.message:'无法读取空间详情');}).finally(()=>{if(alive)setLoading(false);});return()=>{alive=false;};},[id,refresh]);
 const changed=async()=>{setPending(undefined);setConfirmName('');setRefresh(value=>value+1);setNotice('空间管理操作已保存');await onChanged();};
 const confirm=(value:Pending)=>{setPending(value);setConfirmName('');setNotice('');};
 const execute=()=>run(async()=>{if(!pending)return;switch(pending.kind){
  case 'delete':await cloudJson(path('/admin/workspace',id),'DELETE',{confirmName});await onChanged();onBack();return;
  case 'owner':await cloudJson(path('/admin/workspace-owner',id),'POST',{user:pending.user.id,confirmName});break;
  case 'remove':await cloudJson(path('/admin/workspace-member',id),'DELETE',{user:pending.user.id});break;
  case 'member':await cloudJson(path('/admin/workspace-member',id),pending.add?'PUT':'PATCH',{user:pending.user.id,role:pending.role});break;
 }await changed();});
 return <section className="server-workspace-settings"><button className="workspace-back" disabled={busy} onClick={onBack}>返回所有团队空间</button>
  {error&&<p role="alert" className="form-error">{error}<button disabled={busy} onClick={()=>setRefresh(value=>value+1)}>重试</button></p>}{notice&&<p role="status" className="form-notice">{notice}</p>}
  {loading?<p role="status">正在读取空间详情…</p>:detail&&<><h2>{detail.name}</h2><p className="organization-help">{detail.members.length} 位成员 · {detail.documents} 篇文档 · 存储占用 {(detail.storageBytes/1024/1024).toFixed(2)} MB</p>
   <form className="global-space-toolbar" onSubmit={event=>{event.preventDefault();run(async()=>{await cloudJson(path('/admin/workspace',id),'PATCH',{name});await changed();});}}><label className="form-field">空间名称<input required maxLength={160} value={name} disabled={busy||!!pending} onChange={event=>setName(event.target.value)}/></label><button disabled={busy||!!pending||!name.trim()||name.trim()===detail.name}>保存名称</button></form>
   <h3>空间成员</h3><ul className="cloud-members global-space-list" aria-label="全局管理空间成员">{detail.members.map(member=><li key={member.id}><div><strong>{member.name}</strong><small>{member.email}</small><small>{roles[member.role]}</small></div><div className="global-space-actions">{member.role!=='owner'&&<><select aria-label={member.name+'的角色'} value={member.role} disabled={busy||!!pending} onChange={event=>confirm({kind:'member',user:member,role:event.target.value as Role,add:false})}>{(['admin','member','viewer'] as const).map(role=><option key={role} value={role}>{roles[role]}</option>)}</select><button disabled={busy||!!pending} onClick={()=>confirm({kind:'owner',user:member})}>设为所有者</button><button className="danger" disabled={busy||!!pending} onClick={()=>confirm({kind:'remove',user:member})}>移除</button></>}</div></li>)}</ul>
   <AdminUserSearch busy={busy||!!pending} members={detail.members} onAdd={user=>confirm({kind:'member',user,role:'member',add:true})} onOwner={user=>confirm({kind:'owner',user})}/>
   <section className="workspace-panel"><h3>空间维护</h3><p className="organization-help">删除会永久移除空间中的文档和附件，操作记录保留。</p><button className="danger" disabled={busy||!!pending} onClick={()=>confirm({kind:'delete'})}>删除工作空间</button></section>
   {pending&&<div className="cloud-confirm" role="alertdialog" aria-label="确认全局空间操作"><h3>确认{pending.kind==='delete'?'删除空间':pending.kind==='owner'?'重新指定所有者':pending.kind==='remove'?'移除成员':pending.add?'添加成员':'修改角色'}</h3><p>空间：{detail.name}</p>{pending.kind!=='delete'&&<p>用户：{pending.user.name} · {pending.user.email}{pending.kind==='member'?'，角色：'+roles[pending.role]:''}</p>}{pending.kind==='owner'&&<p>原所有者将变为管理员。此用户将获得空间所有者权限。</p>}{(pending.kind==='delete'||pending.kind==='owner')&&<label className="form-field">输入空间名称“{detail.name}”以确认<input autoFocus value={confirmName} disabled={busy} onChange={event=>setConfirmName(event.target.value)}/></label>}<div className="dialog-actions"><button disabled={busy} onClick={()=>setPending(undefined)}>取消</button><button className={pending.kind==='delete'||pending.kind==='remove'?'danger':'primary'} disabled={busy||((pending.kind==='delete'||pending.kind==='owner')&&confirmName!==detail.name)} onClick={execute}>确认执行</button></div></div>}
  </>}
  <AuditHistory key={refresh} id={id} busy={busy} refresh={refresh}/>
 </section>;
}
function AdminUserSearch({busy,members,onAdd,onOwner}:{busy:boolean;members:Member[];onAdd:(user:Member)=>void;onOwner:(user:Member)=>void}){
 const [query,setQuery]=useState(''),[users,setUsers]=useState<Member[]>(),[offset,setOffset]=useState(0),[total,setTotal]=useState(0),[search,setSearch]=useState<string>(),[loading,setLoading]=useState(false),[error,setError]=useState(''),[refresh,setRefresh]=useState(0);
 useEffect(()=>{if(search===undefined)return;let alive=true;setLoading(true);setError('');setUsers(undefined);void cloudJson<{users:Member[];total:number}>('/admin/users?q='+encodeURIComponent(search)+'&offset='+offset).then(value=>{if(alive){setUsers(value.users);setTotal(value.total);}}).catch(e=>{if(alive)setError(e instanceof Error?e.message:'无法搜索用户');}).finally(()=>{if(alive)setLoading(false);});return()=>{alive=false;};},[search,offset,refresh]);
 return <section className="workspace-panel"><h3>添加成员或指定新所有者</h3><form className="global-space-toolbar" onSubmit={event=>{event.preventDefault();setOffset(0);setSearch(query.trim());setRefresh(value=>value+1);}}><label className="form-field">搜索服务器用户<input maxLength={254} value={query} disabled={busy} onChange={event=>setQuery(event.target.value)}/></label><button disabled={busy||loading}>搜索用户</button></form>{error&&<p role="alert">{error}<button disabled={busy} onClick={()=>setRefresh(value=>value+1)}>重试</button></p>}{loading&&<p role="status">正在搜索用户…</p>}{users&&<><ul className="cloud-members global-space-list">{users.map(user=>{const member=members.find(member=>member.id===user.id);return <li key={user.id}><div><strong>{user.name}</strong><small>{user.email}</small></div><div className="global-space-actions"><button disabled={busy||!!member} onClick={()=>onAdd(user)}>{member?'已加入':'添加为成员'}</button><button disabled={busy||member?.role==='owner'} onClick={()=>onOwner(user)}>设为所有者</button></div></li>;})}</ul>{!users.length&&<p>没有符合条件的用户，请调整搜索条件。</p>}<Pagination offset={offset} total={total} limit={100} busy={busy||loading} change={setOffset}/></>}</section>;
}

function AuditHistory({id,busy,refresh}:{id?:string;busy:boolean;refresh:number}){
 const [audit,setAudit]=useState<AuditPage>(),[offset,setOffset]=useState(0),[error,setError]=useState(''),[retry,setRetry]=useState(0);
 useEffect(()=>{let alive=true;setAudit(undefined);setError('');void cloudJson<AuditPage>('/admin/workspace-audit?offset='+offset+(id?'&workspace='+encodeURIComponent(id):'')).then(value=>{if(alive)setAudit(value);}).catch(e=>{if(alive)setError(e instanceof Error?e.message:'无法读取操作记录');});return()=>{alive=false;};},[id,offset,refresh,retry]);
 return <section className="workspace-panel"><h3>全局管理操作记录</h3>{!id&&<p className="organization-help">包含已删除空间的管理记录。</p>}{error?<p role="alert">{error}<button disabled={busy} onClick={()=>setRetry(value=>value+1)}>重试</button></p>:!audit?<p role="status">正在读取操作记录…</p>:<><ul className="cloud-members global-space-list">{audit.entries.map(entry=><li key={entry.id}><div><strong>{entry.details}</strong><small>{entry.actorName} · {date(entry.createdAt)}</small></div></li>)}</ul>{!audit.entries.length&&<p>暂无全局管理操作记录。</p>}<Pagination offset={offset} total={audit.total} limit={audit.limit} busy={busy} change={setOffset}/></>}</section>;
}
