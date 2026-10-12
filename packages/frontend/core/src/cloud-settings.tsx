/**
 * 团队工作空间设置页：概览、成员管理、邀请管理与空间维护。
 *
 * 所有破坏性操作都要显式确认；界面按服务端返回的最新角色渲染，
 * 权限变化后刷新概览与列表，避免用过期权限显示可用操作。
 */
import {useCallback,useEffect,useMemo,useRef,useState} from 'react';
import {cloudJson,type CloudAccount,type CloudWorkspace} from '../../../common/nbstore/src/cloud';

export interface Member {id:string;email:string;name:string;role:CloudWorkspace['role']}
export interface Invite {id:string;email:string;role:string;expiresAt:number;expired:boolean}
export interface Overview {
  id:string;name:string;createdAt:number;role:CloudWorkspace['role'];
  owner:{name:string;email:string}|null;
  members:{total:number;owners:number;admins:number;members:number;viewers:number};
  documents:number;invites:number;
  storage:{documents:number;assets:number;assetCount:number;total:number};
}
export const roleLabels:Record<string,string>={owner:'所有者',admin:'管理员',member:'成员',viewer:'查看者'};
const PAGE_SIZE=20;
export function formatBytes(bytes:number){
  if(bytes<1024)return `${bytes} B`;
  if(bytes<1024*1024)return `${(bytes/1024).toFixed(1)} KB`;
  if(bytes<1024*1024*1024)return `${(bytes/1024/1024).toFixed(1)} MB`;
  return `${(bytes/1024/1024/1024).toFixed(2)} GB`;
}

type Pending=
 |{kind:'role';member:Member;role:CloudWorkspace['role']}
 |{kind:'remove';member:Member}
 |{kind:'leave'}
 |{kind:'transfer';member:Member}
 |{kind:'delete'};

export function WorkspaceSettings({workspace,account,busy,onBack,onChanged,onLeft,onDeleted,run}:{
 workspace:CloudWorkspace;account:CloudAccount;busy:boolean;
 onBack:()=>void;
 /** 名称或成员变化后刷新外层工作空间列表。 */
 onChanged:(next?:Partial<CloudWorkspace>)=>Promise<void>;
 onLeft:()=>void;onDeleted:()=>void;
 run:(work:()=>Promise<void>)=>void;
}){
 const [overview,setOverview]=useState<Overview>(),[members,setMembers]=useState<Member[]>([]),[invites,setInvites]=useState<Invite[]>([]);
 const [search,setSearch]=useState(''),[roleFilter,setRoleFilter]=useState(''),[page,setPage]=useState(0);
 const [pending,setPending]=useState<Pending|null>(null),[token,setToken]=useState(''),[notice,setNotice]=useState('');
 const alive=useRef(true);
 useEffect(()=>{alive.current=true;return()=>{alive.current=false;};},[]);
 const path=useCallback((route:string)=>route+'?workspace='+encodeURIComponent(workspace.id),[workspace.id]);
 // 概览包含服务端最新角色；管理接口按该角色决定是否请求邀请列表。
 const refresh=useCallback(async()=>{
  const next=await cloudJson<Overview>(path('/overview'));
  const list=await cloudJson<Member[]>(path('/members'));
  const manageable=['owner','admin'].includes(next.role);
  const pendingInvites=manageable?await cloudJson<Invite[]>(path('/invites')):[];
  if(!alive.current)return;
  setOverview(next);setMembers(list);setInvites(pendingInvites);
  if(!manageable)setToken('');
 },[path]);
 useEffect(()=>{run(refresh);},[refresh]);
 // 其他管理页面或成员变化后保持同步，避免长时间停留在过期权限上。
 useEffect(()=>{const timer=setInterval(()=>{if(!busy&&!pending)void refresh().catch(()=>{});},8000);return()=>clearInterval(timer);},[refresh,busy,pending]);
 const role=overview?.role??workspace.role;
 const canManage=['owner','admin'].includes(role);
 const isOwner=role==='owner';
 const filtered=useMemo(()=>{
  const keyword=search.trim().toLocaleLowerCase();
  return members.filter(member=>(!roleFilter||member.role===roleFilter)&&(!keyword||member.name.toLocaleLowerCase().includes(keyword)||member.email.toLocaleLowerCase().includes(keyword)))
   .sort((a,b)=>['owner','admin','member','viewer'].indexOf(a.role)-['owner','admin','member','viewer'].indexOf(b.role)||a.name.localeCompare(b.name));
 },[members,search,roleFilter]);
 const pages=Math.max(1,Math.ceil(filtered.length/PAGE_SIZE));
 const current=Math.min(page,pages-1);
 const visible=filtered.slice(current*PAGE_SIZE,current*PAGE_SIZE+PAGE_SIZE);
 useEffect(()=>{setPage(0);},[search,roleFilter]);
 // 管理员不能操作所有者或其他管理员；界面与服务端校验保持一致。
 const canModify=(member:Member)=>canManage&&member.id!==account.id&&member.role!=='owner'&&(isOwner||member.role!=='admin');
 const act=(work:()=>Promise<void>,message:string)=>run(async()=>{await work();await refresh();if(alive.current){setPending(null);setNotice(message);}});
 return <section className="cloud-settings" aria-label="工作空间设置">
  <div className="cloud-settings-head"><button disabled={busy} onClick={onBack}>返回工作空间列表</button><h3>{overview?.name??workspace.name}</h3></div>
  {notice&&<p role="status" className="cloud-settings-notice">{notice}<button aria-label="关闭提示" onClick={()=>setNotice('')}>×</button></p>}
  {!overview?<p role="status">正在读取工作空间信息…</p>:<>
   <dl className="cloud-overview" aria-label="工作空间概览">
    <div><dt>名称</dt><dd>{overview.name}</dd></div>
    <div><dt>所有者</dt><dd>{overview.owner?`${overview.owner.name} · ${overview.owner.email}`:'未知'}</dd></div>
    <div><dt>我的角色</dt><dd>{roleLabels[overview.role]}</dd></div>
    <div><dt>成员数量</dt><dd>{overview.members.total} 人（管理员 {overview.members.admins} · 成员 {overview.members.members} · 查看者 {overview.members.viewers}）</dd></div>
    <div><dt>文档数量</dt><dd>{overview.documents} 篇</dd></div>
    <div><dt>存储使用</dt><dd>{formatBytes(overview.storage.total)}（文档 {formatBytes(overview.storage.documents)} · 附件 {overview.storage.assetCount} 个 {formatBytes(overview.storage.assets)}）</dd></div>
    <div><dt>创建时间</dt><dd>{new Date(overview.createdAt).toLocaleString()}</dd></div>
   </dl>
   {canManage&&<form className="cloud-settings-form" onSubmit={event=>{event.preventDefault();const name=String(new FormData(event.currentTarget).get('name'));act(async()=>{await cloudJson(path('/workspace'),'PATCH',{name});await onChanged({name});},'工作空间名称已更新');}}>
    <label className="form-field">工作空间名称<input name="name" defaultValue={overview.name} required maxLength={160} disabled={busy}/></label>
    <button disabled={busy}>保存名称</button>
   </form>}

   <h3>成员</h3>
   <div className="cloud-filters">
    <label className="form-field">搜索成员<input type="search" value={search} placeholder="姓名或邮箱" disabled={busy} onChange={event=>setSearch(event.target.value)}/></label>
    <label className="form-field">角色筛选<select aria-label="角色筛选" value={roleFilter} disabled={busy} onChange={event=>setRoleFilter(event.target.value)}><option value="">全部角色</option>{Object.entries(roleLabels).map(([value,label])=><option key={value} value={value}>{label}</option>)}</select></label>
    <span role="status">共 {members.length} 人 · 符合条件 {filtered.length} 人</span>
   </div>
   {!filtered.length?<p className="panel-empty">{members.length?'没有符合条件的成员，可清除搜索或筛选。':'这个工作空间还没有其他成员。'}</p>:<>
    <ul className="cloud-members">{visible.map(member=><li key={member.id}>
     <div><strong>{member.name}{member.id===account.id&&'（我）'}</strong><small>{member.email}</small></div>
     <span>{roleLabels[member.role]}</span>
     {isOwner&&member.role!=='owner'&&member.id!==account.id&&<label className="form-field"><span className="sr-only">{member.name}的角色</span>
      <select aria-label={member.name+'的角色'} value={member.role} disabled={busy} onChange={event=>setPending({kind:'role',member,role:event.target.value as CloudWorkspace['role']})}>
       {['admin','member','viewer'].map(value=><option key={value} value={value}>{roleLabels[value]}</option>)}
      </select></label>}
     {canModify(member)&&<button disabled={busy} onClick={()=>setPending({kind:'remove',member})}>移除<span className="sr-only"> {member.name}</span></button>}
     {isOwner&&member.id!==account.id&&member.role!=='owner'&&<button disabled={busy} onClick={()=>setPending({kind:'transfer',member})}>转移所有权<span className="sr-only"> 给 {member.name}</span></button>}
    </li>)}</ul>
    {pages>1&&<div className="cloud-pagination" role="group" aria-label="成员分页">
     <button disabled={busy||current===0} onClick={()=>setPage(current-1)}>上一页</button>
     <span role="status">第 {current+1} / {pages} 页</span>
     <button disabled={busy||current>=pages-1} onClick={()=>setPage(current+1)}>下一页</button>
    </div>}
   </>}

   {canManage&&<>
    <h3>邀请</h3>
    {!invites.length?<p className="panel-empty">没有待接受的邀请。</p>:<ul className="cloud-members" aria-label="邀请列表">{invites.map(invite=><li key={invite.id}>
     <div><strong>{invite.email}</strong><small>{roleLabels[invite.role]??invite.role} · {invite.expired?'已过期，需重新生成':`有效至 ${new Date(invite.expiresAt).toLocaleString()}`}</small></div>
     <span>{invite.expired?'已过期':'待接受'}</span>
     <button disabled={busy} onClick={()=>act(async()=>{await cloudJson(path('/invite'),'DELETE',{id:invite.id});},'邀请已撤销，原口令立即失效')}>撤销<span className="sr-only"> {invite.email}</span></button>
    </li>)}</ul>}
    <form className="cloud-settings-form" onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);act(async()=>{const invite=await cloudJson<{token:string}>(path('/invite'),'POST',{email:data.get('email'),role:data.get('role')});if(alive.current)setToken(invite.token);},'邀请口令已生成，同一邮箱的旧口令已失效');}}>
     <label className="form-field">邀请成员邮箱<input name="email" type="email" required disabled={busy}/></label>
     <label className="form-field">邀请角色<select name="role" disabled={busy}><option value="member">成员（可编辑）</option><option value="viewer">查看者（只读）</option></select></label>
     <button disabled={busy}>生成邀请口令</button>
     {token&&<label className="form-field">邀请口令（7 天有效，仅指定邮箱可接受）<input readOnly value={token} onFocus={event=>event.currentTarget.select()}/></label>}
    </form>
   </>}

   <h3>空间维护</h3>
   {!isOwner&&<button disabled={busy} onClick={()=>setPending({kind:'leave'})}>退出此工作空间</button>}
   {isOwner&&overview.members.total<2&&<p className="organization-help">只有你一个成员时无法转移所有权。先邀请成员，或直接删除工作空间。</p>}
   {isOwner&&<button className="danger" disabled={busy} onClick={()=>setPending({kind:'delete'})}>删除工作空间</button>}

   {pending&&<div className="cloud-confirm" role="alertdialog" aria-label="确认操作">
    {pending.kind==='role'&&<><p>将“{pending.member.name}”的角色改为{roleLabels[pending.role]}？该成员的当前连接会断开并按新权限重新载入。</p>
     <div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setPending(null)}>取消</button><button className="primary" disabled={busy} onClick={()=>act(async()=>{await cloudJson(path('/member'),'PATCH',{user:pending.member.id,role:pending.role});await onChanged();},'成员角色已更新')}>确认调整角色</button></div></>}
    {pending.kind==='remove'&&<><p>移除“{pending.member.name}”？该成员会立即失去访问权限，未提交内容只能由其自行导出草稿。</p>
     <div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setPending(null)}>取消</button><button className="danger" disabled={busy} onClick={()=>act(async()=>{await cloudJson(path('/member'),'DELETE',{user:pending.member.id});await onChanged();},'成员已移除')}>确认移除成员</button></div></>}
    {pending.kind==='leave'&&<><p>退出“{overview.name}”？退出后需要新的邀请口令才能重新加入。</p>
     <div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setPending(null)}>取消</button><button className="danger" disabled={busy} onClick={()=>run(async()=>{await cloudJson(path('/leave'),'POST');await onChanged();if(alive.current){setPending(null);onLeft();}})}>确认退出</button></div></>}
    {pending.kind==='transfer'&&<><p>把所有权转移给“{pending.member.name}”？你会变成管理员，之后不能删除工作空间或修改其他管理员。</p>
     <div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setPending(null)}>取消</button><button className="danger" disabled={busy} onClick={()=>act(async()=>{await cloudJson(path('/transfer'),'POST',{user:pending.member.id});await onChanged();},'所有权已转移，你现在是管理员')}>确认转移所有权</button></div></>}
    {pending.kind==='delete'&&<><p>删除“{overview.name}”将永久移除 {overview.documents} 篇文档、{overview.storage.assetCount} 个附件（共 {formatBytes(overview.storage.total)}）及 {overview.members.total} 名成员的访问权，操作不可撤销。</p>
     <form onSubmit={event=>{event.preventDefault();run(async()=>{await cloudJson(path('/workspace'),'DELETE');await onChanged();if(alive.current){setPending(null);onDeleted();}});}}>
      <label className="form-field">输入工作空间名称确认删除<input required pattern={overview.name.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')} disabled={busy}/></label>
      <div className="cloud-confirm-actions"><button type="button" disabled={busy} onClick={()=>setPending(null)}>取消</button><button className="danger" disabled={busy}>永久删除工作空间及全部内容</button></div>
     </form></>}
   </div>}
  </>}
 </section>;
}
