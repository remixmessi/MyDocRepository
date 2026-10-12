import {useEffect,useState} from 'react';
import {Modal} from './dialogs';
import {cloudJson,cloudServerUrl,configureCloudServer,type CloudAccount,type CloudMailStatus,type CloudWorkspace} from '../../../common/nbstore/src/cloud';
import {projectPreferences} from '@mydoc/nbstore';
import {WorkspaceSettings,roleLabels} from './cloud-settings';
import {AccountSettings} from './account-settings';
import {ServerAdminSettings} from './server-admin-settings';
import {ArrowUpRight,Building2,Plus,Settings2,UserRound,Users} from 'lucide-react';
/**
 * 登录／注册与找回密码。
 * 找回密码不显示邮箱是否存在；未配置邮件服务时明确说明口令只写入服务端日志。
 */
export function CloudLogin({onLogin,beforeSession}:{onLogin:()=>void;beforeSession?:()=>Promise<void>}){
 const [mode,setMode]=useState<'login'|'register'|'forgot'|'reset'>('login'),[busy,setBusy]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState('');
 const register=mode==='register';
 const run=(work:()=>Promise<void>)=>{setBusy(true);setError('');void (async()=>{try{await work();}catch(error){setError(error instanceof Error?error.message:'操作失败，请重试');}finally{setBusy(false);}})();};
 if(mode==='forgot'||mode==='reset')return <div>
  <p className="organization-help">{mode==='forgot'?'输入账户邮箱获取一次性重置口令。为避免暴露账户是否存在，无论邮箱是否注册都会返回相同结果。':'输入收到的重置口令与新密码。口令 1 小时内有效、只能使用一次；重置后全部设备都需要重新登录。'}</p>
  {notice&&<p role="status" className="form-notice">{notice}</p>}
  {mode==='forgot'?<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
   run(async()=>{const result=await cloudJson<{mailConfigured:boolean;notice?:string}>('/forgot','POST',{email:data.get('email')});
    setNotice(result.notice??'如果该邮箱已注册，重置口令已发送，请查收。');setMode('reset');});}}>
   <label className="form-field">账户邮箱<input type="email" name="email" autoComplete="email" required maxLength={254} disabled={busy}/></label>
   {error&&<p role="alert" className="form-error">{error}</p>}
   <div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>{setMode('login');setError('');setNotice('');}}>返回登录</button><button className="primary" disabled={busy}>获取重置口令</button></div>
  </form>:<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
   run(async()=>{await cloudJson('/reset','POST',{token:data.get('token'),password:data.get('password')});setNotice('密码已重置，请用新密码登录。');setMode('login');});}}>
   <label className="form-field">重置口令<input name="token" required maxLength={128} disabled={busy}/></label>
   <label className="form-field">新密码（至少 10 位）<input type="password" name="password" autoComplete="new-password" required minLength={10} maxLength={256} disabled={busy}/></label>
   {error&&<p role="alert" className="form-error">{error}</p>}
   <div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>{setMode('forgot');setError('');}}>重新获取口令</button><button className="primary" disabled={busy}>重置密码</button></div>
  </form>}
 </div>;
 return <form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
  run(async()=>{await (beforeSession?.()??Promise.resolve());await cloudJson(register?'/register':'/login','POST',{email:data.get('email'),password:data.get('password'),name:data.get('name')});onLogin();});}}>
  <p className="organization-help">登录后可创建或加入团队工作空间，与成员共同编辑文档。个人工作区仍保存在本机，不会上传。</p>
  {notice&&<p role="status" className="form-notice">{notice}</p>}
  {register&&<label className="form-field">显示名称<input name="name" required maxLength={160} disabled={busy}/></label>}
  <label className="form-field">邮箱<input type="email" name="email" autoComplete="email" required disabled={busy}/></label>
  <label className="form-field">密码<input type="password" name="password" minLength={10} maxLength={256} autoComplete={register?'new-password':'current-password'} required disabled={busy}/></label>
  {error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions">
   <button type="button" disabled={busy} onClick={()=>{setMode(register?'login':'register');setError('');setNotice('');}}>{register?'已有账户，去登录':'创建账户'}</button>
   {!register&&<button type="button" disabled={busy} onClick={()=>{setMode('forgot');setError('');setNotice('');}}>忘记密码</button>}
   <button className="primary" disabled={busy}>{busy?'正在处理…':register?'注册并登录':'登录'}</button>
  </div>
 </form>;
}
export function CloudDialog({onClose,onRefresh,onOpen}:{onClose:()=>void;onRefresh:()=>Promise<void>;onOpen:(id:string)=>void}){
 const [user,setUser]=useState<CloudAccount|null>(),[spaces,setSpaces]=useState<CloudWorkspace[]>([]),[selected,setSelected]=useState<CloudWorkspace>(),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const [mail,setMail]=useState<CloudMailStatus>({configured:false}),[account,setAccount]=useState(false),[admin,setAdmin]=useState(false);
 const [serverUrl,setServerUrl]=useState(cloudServerUrl()),[serverDraft,setServerDraft]=useState(cloudServerUrl()),[serverNotice,setServerNotice]=useState('');
 const [spaceQuery,setSpaceQuery]=useState(''),[spaceAction,setSpaceAction]=useState<'create'|'join'>('create');
 const isDesktop=!!window.mydocDesktop;
 const load=async()=>{const session=await cloudJson<{user:CloudAccount|null;mail?:CloudMailStatus}>('/session');setUser(session.user);setMail(session.mail??{configured:false});if(session.user)setSpaces(await cloudJson('/workspaces'));};
 const run=(work:()=>Promise<void>)=>{setBusy(true);setError('');void (async()=>{try{await work();}catch(error){setError(error instanceof Error?error.message:'操作失败，请重试');}finally{setBusy(false);}})();};
 useEffect(()=>{let alive=true;void (async()=>{if(isDesktop){const saved=projectPreferences.getItem('cloud-server-url');if(saved){try{configureCloudServer(saved);if(alive){setServerUrl(saved);setServerDraft(saved);}}catch{}}}run(load);})();return()=>{alive=false;};},[]);
 const saveServer=()=>run(async()=>{const next=configureCloudServer(serverDraft);if(isDesktop)await projectPreferences.setItem('cloud-server-url',next);setServerUrl(next);setServerNotice(next?'服务器地址已保存':'已恢复使用当前应用服务器');await load();});
 return <Modal title="团队工作空间" wide onClose={()=>{if(!busy)onClose();}}>{error&&<p role="alert" className="form-error">{error}</p>}{isDesktop&&<form className="cloud-server-config" onSubmit={event=>{event.preventDefault();saveServer();}}><label className="form-field">服务器地址<input type="url" value={serverDraft} placeholder="https://docs.example.com" onChange={event=>setServerDraft(event.target.value)} disabled={busy}/><small>桌面端注册、登录和同步将使用此地址；留空恢复为当前服务器。</small></label><button disabled={busy}>{busy?'正在连接…':'保存并连接'}</button>{serverNotice&&<span role="status">{serverNotice}</span>}{serverUrl&&<small role="status">当前：{serverUrl}</small>}</form>}{user===undefined?<div role="status">{error?<button onClick={()=>run(load)}>重试加载</button>:'正在读取账户…'}</div>:!user?<CloudLogin beforeSession={onRefresh} onLogin={()=>location.reload()}/>:<div className="cloud-management">
  <div className="cloud-account">
   <span className="management-identity"><UserRound size={17}/><span><strong>{user.name}</strong><small>{user.email}</small></span></span>
   {!account&&!selected&&!admin&&<button title="账户设置" aria-label="账户设置" disabled={busy} onClick={()=>setAccount(true)}><Settings2 size={17}/></button>}
   {!account&&!selected&&!admin&&user.serverAdmin&&<button disabled={busy} onClick={()=>setAdmin(true)}>服务器用户管理</button>}
   <button disabled={busy} onClick={()=>run(async()=>{await onRefresh();await cloudJson('/logout','POST');location.reload();})}>退出登录</button>
  </div>
  {admin&&user.serverAdmin?<ServerAdminSettings accountId={user.id} busy={busy} run={run} onBack={()=>setAdmin(false)}/>:account?<AccountSettings account={user} mail={mail} busy={busy} run={run}
    onBack={()=>setAccount(false)}
    onChanged={async()=>{await load();await onRefresh();}}
    onSignedOut={()=>location.reload()}/>
  :selected?<WorkspaceSettings workspace={selected} account={user} busy={busy} run={run}
    onBack={()=>setSelected(undefined)}
    onChanged={async next=>{if(next)setSelected(current=>current?{...current,...next}:current);await load();await onRefresh();}}
    onLeft={()=>setSelected(undefined)} onDeleted={()=>setSelected(undefined)}/>
  :<div className="workspace-directory">
   <header className="workspace-directory-head"><div><h2>工作空间</h2><span>{spaces.length} 个团队空间</span></div><Building2 size={24}/></header>
   <label className="form-field"><span className="sr-only">搜索工作空间</span><input type="search" placeholder="搜索工作空间" value={spaceQuery} onChange={event=>setSpaceQuery(event.target.value)}/></label>
   <ul className="cloud-members workspace-directory-list">{spaces.filter(space=>space.name.toLowerCase().includes(spaceQuery.trim().toLowerCase())).map(space=><li key={space.id}>
    <span className="workspace-monogram" aria-hidden="true">{space.name.slice(0,1)}</span><div><strong>{space.name}</strong><small>{new Date(space.createdAt).toLocaleDateString('zh-CN')}</small></div><span className="workspace-role">{roleLabels[space.role]}</span>
    <button title="打开工作空间" disabled={busy} onClick={()=>run(async()=>{await onRefresh();onOpen(space.id);})}>打开 <ArrowUpRight size={14}/></button>
    <button title="设置与成员" aria-label="设置与成员" disabled={busy} onClick={()=>setSelected(space)}><Settings2 size={16}/></button>
   </li>)}</ul>
   {!spaces.length&&<p className="panel-empty">还没有团队工作空间。</p>}{!!spaces.length&&!spaces.some(space=>space.name.toLowerCase().includes(spaceQuery.trim().toLowerCase()))&&<p className="panel-empty">没有匹配的工作空间。</p>}
   <div className="workspace-action-tabs" role="group" aria-label="工作空间操作"><button aria-pressed={spaceAction==='create'} onClick={()=>setSpaceAction('create')}><Plus size={15}/>创建空间</button><button aria-pressed={spaceAction==='join'} onClick={()=>setSpaceAction('join')}><Users size={15}/>加入空间</button></div>
   {spaceAction==='create'&&<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);run(async()=>{await cloudJson('/workspaces','POST',{name:data.get('name')});await load();await onRefresh();});}}>
    <label className="form-field">新工作空间名称<input name="name" required maxLength={160} disabled={busy}/></label>
    <button className="primary" disabled={busy}>创建团队工作空间</button>
   </form>}
   {spaceAction==='join'&&<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);run(async()=>{await cloudJson('/accept','POST',{token:data.get('token')});await load();await onRefresh();});}}>
    <label className="form-field">邀请口令<input name="token" required maxLength={128} disabled={busy}/></label>
    <button disabled={busy}>接受邀请</button>
   </form>}
  </div>}
 </div>}</Modal>;
}
