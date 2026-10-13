import {useEffect,useRef,useState} from 'react';
import {ArrowLeft} from 'lucide-react';
import {cloudJson,cloudRequest,cloudUrl,type CloudAccount,type CloudMailStatus} from '../../../common/nbstore/src/cloud';

/**
 * 账户设置：显示名称与头像、修改密码、登录设备、邮箱验证与更改邮箱。
 * 权限与口令校验全部在服务端完成，这里只负责输入、确认与把服务端结果如实显示。
 */
export interface Session {id:string;createdAt:number;lastSeenAt:number;expiresAt:number;agent:string|null;ip:string|null;current:boolean}
const AVATAR_TYPES='image/png,image/jpeg,image/webp',AVATAR_LIMIT=256*1024;
export const formatTime=(value:number)=>value?new Date(value).toLocaleString('zh-CN',{hour12:false}):'—';
/** User-Agent 只做粗略归类：用于区分设备，不做精确识别。 */
export function describeAgent(agent:string|null){
 if(!agent)return '未知设备';
 const platform=/Windows/i.test(agent)?'Windows':/Macintosh|Mac OS/i.test(agent)?'macOS':/iPhone|iPad/i.test(agent)?'iOS':/Android/i.test(agent)?'Android':/Linux/i.test(agent)?'Linux':'未知系统';
 const browser=/Edg\//i.test(agent)?'Edge':/OPR\//i.test(agent)?'Opera':/Chrome\//i.test(agent)?'Chrome':/Safari\//i.test(agent)?'Safari':/Firefox\//i.test(agent)?'Firefox':agent.split(/[\s/]/)[0]||'未知客户端';
 return platform+' · '+browser;
}

export function AccountSettings({account,mail,busy,run,onBack,onChanged,onSignedOut}:{account:CloudAccount;mail:CloudMailStatus;busy:boolean;run:(work:()=>Promise<void>)=>void;onBack:()=>void;onChanged:()=>Promise<void>;onSignedOut:()=>void}){
 const [sessions,setSessions]=useState<Session[]>([]),[notice,setNotice]=useState(''),[avatarTick,setAvatarTick]=useState(0);
 const [verifyToken,setVerifyToken]=useState(''),[emailToken,setEmailToken]=useState(''),[pendingEmail,setPendingEmail]=useState('');
 const [confirm,setConfirm]=useState<{kind:'session';session:Session}|{kind:'others'}|null>(null);
 const alive=useRef(true);useEffect(()=>()=>{alive.current=false;},[]);
 const loadSessions=async()=>{const rows=await cloudJson<Session[]>('/sessions');if(alive.current)setSessions(rows);};
 useEffect(()=>{run(loadSessions);const timer=setInterval(()=>void loadSessions().catch(()=>{}),8000);return()=>clearInterval(timer);},[account.id]);
 const act=(work:()=>Promise<void>,message:string)=>run(async()=>{await work();if(alive.current){setNotice(message);setConfirm(null);}});

 return <section className="account-settings">
  <header className="cloud-settings-head"><button className="workspace-back" disabled={busy} onClick={onBack}><ArrowLeft size={16}/>返回工作空间列表</button><h3>账户设置</h3></header>
  {notice&&<p role="status" className="form-notice">{notice}</p>}
  {!mail.configured&&<p className="organization-help">邮件服务未配置（{mail.reason??'缺少 SMTP 设置'}）。验证与更改邮箱的一次性口令会直接显示在此页面；找回密码的口令只写入服务端日志，需要联系运维获取。</p>}

  <h4>个人资料</h4>
  <dl className="cloud-overview" aria-label="账户概览">
   <div><dt>邮箱</dt><dd>{account.email}{account.verified?' · 已验证':' · 未验证'}</dd></div>
   <div><dt>注册时间</dt><dd>{formatTime(account.createdAt??0)}</dd></div>
   <div><dt>当前会话到期</dt><dd>{formatTime(account.expiresAt??0)}</dd></div>
  </dl>
  <form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);act(async()=>{await cloudJson('/profile','PATCH',{name:data.get('name')});await onChanged();},'显示名称已更新');}}>
   <label className="form-field">显示名称<input name="name" defaultValue={account.name} required maxLength={160} disabled={busy}/></label>
   <button className="primary" disabled={busy}>保存显示名称</button>
  </form>

  <h4>头像</h4>
  <div className="account-avatar">
   {account.hasAvatar?<img src={cloudUrl('/avatar?user='+encodeURIComponent(account.id)+'&v='+avatarTick)} width={64} height={64} alt={account.name+' 的头像'}/>:<span className="account-avatar-empty" aria-hidden="true">{account.name.slice(0,1)}</span>}
   <div>
    <label className="form-field">上传头像（PNG／JPEG／WebP，不超过 256 KiB）
     <input type="file" accept={AVATAR_TYPES} disabled={busy} onChange={event=>{
      const file=event.target.files?.[0];event.target.value='';if(!file)return;
      if(file.size>AVATAR_LIMIT){setNotice('');run(async()=>{throw new Error('头像超过 256 KiB，请先压缩');});return;}
      act(async()=>{await cloudRequest('/avatar',{method:'PUT',headers:{'Content-Type':file.type},body:file});await onChanged();if(alive.current)setAvatarTick(tick=>tick+1);},'头像已更新');
     }}/></label>
    {account.hasAvatar&&<button disabled={busy} onClick={()=>act(async()=>{await cloudJson('/avatar','DELETE');await onChanged();if(alive.current)setAvatarTick(tick=>tick+1);},'头像已移除')}>移除头像</button>}
   </div>
  </div>

  <h4>邮箱验证</h4>
  {account.verified?<p className="organization-help">邮箱已验证。</p>:<>
   <p className="organization-help">验证后才能确保找回密码与通知送达本人。口令 1 小时内有效，只能使用一次。</p>
   <button disabled={busy} onClick={()=>run(async()=>{const result=await cloudJson<{sent:boolean;token?:string;notice?:string}>('/verify','POST',{});if(!alive.current)return;setVerifyToken(result.token??'');setNotice(result.sent?'验证邮件已发送，请查收':result.notice??'口令已生成');})}>{verifyToken?'重新发送验证口令':'发送验证口令'}</button>
   {verifyToken&&<label className="form-field">验证口令（仅本次显示）<input readOnly value={verifyToken} onFocus={event=>event.currentTarget.select()}/></label>}
   <form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);act(async()=>{await cloudJson('/verify','PATCH',{token:data.get('token')});await onChanged();if(alive.current)setVerifyToken('');},'邮箱已验证');}}>
    <label className="form-field">输入收到的验证口令<input name="token" required maxLength={128} disabled={busy}/></label>
    <button className="primary" disabled={busy}>完成验证</button>
   </form>
  </>}

  <h4>修改密码</h4>
  <form onSubmit={event=>{event.preventDefault();const form=event.currentTarget,data=new FormData(form);
   act(async()=>{const result=await cloudJson<{signedOutOthers:boolean}>('/password','POST',{current:data.get('current'),password:data.get('password'),signOutOthers:data.get('others')==='on'});form.reset();await loadSessions();setNotice(result.signedOutOthers?'密码已更新，其他设备需重新登录':'密码已更新，其他设备仍保持登录');},'密码已更新');}}>
   <label className="form-field">当前密码<input type="password" name="current" autoComplete="current-password" required minLength={10} maxLength={256} disabled={busy}/></label>
   <label className="form-field">新密码（至少 10 位）<input type="password" name="password" autoComplete="new-password" required minLength={10} maxLength={256} disabled={busy}/></label>
   <label className="form-check"><input type="checkbox" name="others" defaultChecked disabled={busy}/>同时让其他设备重新登录</label>
   <button className="primary" disabled={busy}>保存新密码</button>
  </form>

  <h4>更改邮箱</h4>
  <p className="organization-help">确认前邮箱不会变更。更改后发往旧地址的未接受邀请将失效，需要邀请人按新邮箱重新邀请。</p>
  <form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
   run(async()=>{const result=await cloudJson<{sent:boolean;email:string;pendingInvites:number;token?:string;notice?:string}>('/email','POST',{email:data.get('email'),password:data.get('password')});
    if(!alive.current)return;setPendingEmail(result.email);setEmailToken(result.token??'');
    setNotice((result.sent?'验证口令已发送到 '+result.email:result.notice??'口令已生成')+(result.pendingInvites?`；旧邮箱还有 ${result.pendingInvites} 个未接受的邀请，更改后将失效`:''));});}}>
   <label className="form-field">新邮箱<input type="email" name="email" required maxLength={254} disabled={busy}/></label>
   <label className="form-field">当前密码（确认是本人操作）<input type="password" name="password" autoComplete="current-password" required minLength={10} maxLength={256} disabled={busy}/></label>
   <button disabled={busy}>发送更改邮箱口令</button>
  </form>
  {emailToken&&<label className="form-field">更改邮箱口令（仅本次显示）<input readOnly value={emailToken} onFocus={event=>event.currentTarget.select()}/></label>}
  {pendingEmail&&<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
   act(async()=>{const result=await cloudJson<{email:string;staleInvites:number}>('/email','PATCH',{token:data.get('token')});await onChanged();
    if(alive.current){setPendingEmail('');setEmailToken('');setNotice('邮箱已更改为 '+result.email+(result.staleInvites?`；旧邮箱的 ${result.staleInvites} 个邀请已失效`:''));}},'邮箱已更改');}}>
   <label className="form-field">输入发往 {pendingEmail} 的口令<input name="token" required maxLength={128} disabled={busy}/></label>
   <button className="primary" disabled={busy}>确认更改邮箱</button>
  </form>}

  <h4>登录设备</h4>
  <ul className="cloud-members" aria-label="登录设备列表">{sessions.map(session=><li key={session.id}>
   <div><strong>{describeAgent(session.agent)}{session.current?'（当前设备）':''}</strong><small>最近活动 {formatTime(session.lastSeenAt)} · 登录于 {formatTime(session.createdAt)} · 到期 {formatTime(session.expiresAt)}{session.ip?' · '+session.ip:''}</small></div>
   {session.current?<button disabled={busy} onClick={()=>run(async()=>{await cloudJson('/logout','POST');onSignedOut();})}>退出登录</button>
    :<button disabled={busy} onClick={()=>setConfirm({kind:'session',session})}>注销<span className="sr-only"> {describeAgent(session.agent)}</span></button>}
  </li>)}</ul>
  {!sessions.length&&<p className="panel-empty">正在读取登录设备…</p>}
  {sessions.length>1&&<button disabled={busy} onClick={()=>setConfirm({kind:'others'})}>注销其他全部设备</button>}

  {confirm&&<div className="cloud-confirm" role="alertdialog" aria-label="确认操作">
   {confirm.kind==='session'&&<><p>注销“{describeAgent(confirm.session.agent)}”？该设备会立即失去访问权限，未提交内容只能在该设备上导出草稿。</p>
    <div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setConfirm(null)}>取消</button><button className="danger" disabled={busy} onClick={()=>act(async()=>{await cloudJson('/sessions','DELETE',{session:confirm.session.id});await loadSessions();},'该设备已注销')}>确认注销设备</button></div></>}
   {confirm.kind==='others'&&<><p>注销除当前设备外的 {sessions.length-1} 个会话？这些设备需要重新登录，未提交内容只能在各自设备上导出草稿。</p>
    <div className="cloud-confirm-actions"><button disabled={busy} onClick={()=>setConfirm(null)}>取消</button><button className="danger" disabled={busy} onClick={()=>act(async()=>{const result=await cloudJson<{removed:number}>('/sessions','DELETE',{all:true});await loadSessions();setNotice('已注销 '+result.removed+' 个其他设备会话');},'已注销其他设备')}>确认注销其他设备</button></div></>}
  </div>}
 </section>;
}
