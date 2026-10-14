import {useEffect,useRef,useState} from 'react';
import {Modal} from './dialogs';
import {ServerConnections} from './server-connections';
import {localAccount,registerLocalOnServer,restoreLocalAccount} from './local-account-ui';
import {cloudJson,cloudServerUrl,type CloudAccount,type CloudMailStatus,type CloudWorkspace} from '../../../common/nbstore/src/cloud';
import {WorkspaceSettings,roleLabels} from './cloud-settings';
import {ArrowUpRight,Building2,Plus,Settings2,Users} from 'lucide-react';
/**
 * 登录／注册与找回密码。
 * 找回密码不显示邮箱是否存在；未配置邮件服务时明确说明口令只写入服务端日志。
 */
export function CloudLogin({onLogin,beforeSession,requestJson=cloudJson}:{onLogin:()=>void|Promise<void>;beforeSession?:()=>Promise<void>;requestJson?:typeof cloudJson}){
 const [mode,setMode]=useState<'login'|'register'|'forgot'|'reset'>('login'),[busy,setBusy]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState('');
 const register=mode==='register';
 const run=(work:()=>Promise<void>)=>{setBusy(true);setError('');void (async()=>{try{await work();}catch(error){setError(error instanceof Error?error.message:'操作失败，请重试');}finally{setBusy(false);}})();};
 if(mode==='forgot'||mode==='reset')return <div>
  <p className="organization-help">{mode==='forgot'?'输入账户邮箱获取一次性重置口令。为避免暴露账户是否存在，无论邮箱是否注册都会返回相同结果。':'输入收到的重置口令与新密码。口令 1 小时内有效、只能使用一次；重置后全部设备都需要重新登录。'}</p>
  {notice&&<p role="status" className="form-notice">{notice}</p>}
  {mode==='forgot'?<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
   run(async()=>{const result=await requestJson<{mailConfigured:boolean;notice?:string}>('/forgot','POST',{email:data.get('email')});
    setNotice(result.notice??'如果该邮箱已注册，重置口令已发送，请查收。');setMode('reset');});}}>
   <label className="form-field">账户邮箱<input type="email" name="email" autoComplete="email" required maxLength={254} disabled={busy}/></label>
   {error&&<p role="alert" className="form-error">{error}</p>}
   <div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>{setMode('login');setError('');setNotice('');}}>返回登录</button><button className="primary" disabled={busy}>获取重置口令</button></div>
  </form>:<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
   run(async()=>{await requestJson('/reset','POST',{token:data.get('token'),password:data.get('password')});setNotice('密码已重置，请用新密码登录。');setMode('login');});}}>
   <label className="form-field">重置口令<input name="token" required maxLength={128} disabled={busy}/></label>
   <label className="form-field">新密码（至少 10 位）<input type="password" name="password" autoComplete="new-password" required minLength={10} maxLength={256} disabled={busy}/></label>
   {error&&<p role="alert" className="form-error">{error}</p>}
   <div className="dialog-actions"><button type="button" disabled={busy} onClick={()=>{setMode('forgot');setError('');}}>重新获取口令</button><button className="primary" disabled={busy}>重置密码</button></div>
  </form>}
 </div>;
 return <form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);
  run(async()=>{await (beforeSession?.()??Promise.resolve());await requestJson(register?'/register':'/login','POST',{email:data.get('email'),password:data.get('password'),name:data.get('name')});await onLogin();});}}>
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
export function CloudDialog({onClose,onRefresh,onOpen,onPrepareSwitch}:{onPrepareSwitch:()=>Promise<void>;onClose:()=>void;onRefresh:()=>Promise<void>;onOpen:(id:string)=>void}){
 const [user,setUser]=useState<CloudAccount|null>(),[spaces,setSpaces]=useState<CloudWorkspace[]>([]),[selected,setSelected]=useState<CloudWorkspace>(),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const noRemoteServer=!!window.mydocDesktop&&!cloudServerUrl();
 const [serverRevision,setServerRevision]=useState(0);
 const requestGeneration=useRef(0);
 const [,refreshLocal]=useState(0);
 useEffect(()=>{let alive=true;void restoreLocalAccount().then(()=>{if(alive)refreshLocal(value=>value+1);}).catch(()=>{});return()=>{alive=false;};},[]);
 const [spaceQuery,setSpaceQuery]=useState(''),[spaceAction,setSpaceAction]=useState<'create'|'join'>('create');
 const load=async(generation=requestGeneration.current)=>{if(noRemoteServer){setUser(null);setSpaces([]);return;}const session=await cloudJson<{user:CloudAccount|null;mail?:CloudMailStatus}>('/session');if(generation!==requestGeneration.current)return;setUser(session.user);if(session.user){const nextSpaces=await cloudJson<CloudWorkspace[]>('/workspaces');if(generation!==requestGeneration.current)return;setSpaces(nextSpaces);}else setSpaces([]);};
 const run=(work:()=>Promise<void>)=>{const generation=++requestGeneration.current;setBusy(true);setError('');void (async()=>{try{await work();if(generation===requestGeneration.current)setError('');}catch(error){if(generation===requestGeneration.current)setError(error instanceof Error?error.message:'操作失败，请重试');}finally{if(generation===requestGeneration.current)setBusy(false);}})();};
 useEffect(()=>{run(load);},[]);
 useEffect(()=>{if(serverRevision)run(load);},[serverRevision]);
 const serverChanged=async()=>{await onRefresh();setServerRevision(value=>value+1);};
 return <Modal title="团队工作空间" wide className="workspace-settings-dialog" onClose={()=>{if(!busy)onClose();}}>
  <div className="team-workspace-dialog">
   <header className="team-workspace-header">
    <div><p className="workspace-eyebrow">TEAM WORKSPACE</p><h2>团队工作空间</h2><p>连接远程服务器后，在这里查看、创建和管理团队空间。</p></div>
    <span className={`team-connection-state ${noRemoteServer?'is-local':'is-connected'}`} role="status">{noRemoteServer?'未连接远程服务器':user?'已登录':'已连接服务器'}</span>
   </header>
   {error&&<p role="alert" className="form-error team-workspace-error">{error}<button type="button" onClick={()=>run(load)}>重试</button></p>}
   <ServerConnections busy={busy} prepare={onPrepareSwitch} onServerChange={serverChanged}/>
   <div className="team-workspace-content">
    {noRemoteServer?<div className="team-workspace-empty" role="status"><Building2 size={22}/><strong>先添加一个远程服务器</strong><span>桌面端不会自动连接本地内部服务。添加地址后，登录和空间列表会在当前窗口继续。</span></div>:user===undefined?<div className="team-workspace-empty" role="status">正在读取服务器账号…</div>:!user?<div className="team-login-panel"><div><h3>登录服务器账号</h3><p className="organization-help">登录后即可查看你加入的团队空间。登录过程会留在当前窗口。</p></div><CloudLogin beforeSession={onRefresh} onLogin={()=>load()}/>{localAccount&&<form className="local-register-form" onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);run(async()=>{await onRefresh();await registerLocalOnServer(String(data.get('password')));await load();});}}><h3>使用本地账号注册服务器</h3><p>{localAccount.name} · {localAccount.email}</p><label className="form-field">确认本地密码<input type="password" name="password" required minLength={10} disabled={busy} autoComplete="current-password"/></label><button disabled={busy}>注册到当前服务器</button></form>}</div>:<div className="team-management">
     {selected?<WorkspaceSettings workspace={selected} account={user} busy={busy} run={run}
       onBack={()=>setSelected(undefined)}
       onChanged={async next=>{if(next)setSelected(current=>current?{...current,...next}:current);await load();await onRefresh();}}
       onLeft={()=>setSelected(undefined)} onDeleted={()=>setSelected(undefined)}/>
     :<div className="workspace-directory">
       <header className="workspace-directory-head"><div><p className="workspace-eyebrow">YOUR SPACES</p><h2>你的团队空间</h2><span>{spaces.length} 个空间</span></div><Building2 size={24}/></header>
       <div className="workspace-directory-toolbar"><label className="form-field"><span className="sr-only">搜索工作空间</span><input type="search" placeholder="搜索空间名称" value={spaceQuery} onChange={event=>setSpaceQuery(event.target.value)}/></label><span className="workspace-refresh-hint">登录状态会在当前窗口同步更新</span></div>
       <ul className="cloud-members workspace-directory-list">{spaces.filter(space=>space.name.toLowerCase().includes(spaceQuery.trim().toLowerCase())).map(space=><li key={space.id}>
        <span className="workspace-monogram" aria-hidden="true">{space.name.slice(0,1)}</span><div><strong>{space.name}</strong><small>{new Date(space.createdAt).toLocaleDateString('zh-CN')}</small></div><span className="workspace-role">{roleLabels[space.role]}</span>
        <button title="打开工作空间" disabled={busy} onClick={()=>run(async()=>{await onRefresh();onOpen(space.id);})}>打开 <ArrowUpRight size={14}/></button>
        <button title="设置与成员" aria-label={`设置与成员：${space.name}`} disabled={busy} onClick={()=>setSelected(space)}><Settings2 size={16}/></button>
       </li>)}</ul>
       {!spaces.length&&<p className="panel-empty">还没有团队工作空间，可以创建一个新的空间。</p>}{!!spaces.length&&!spaces.some(space=>space.name.toLowerCase().includes(spaceQuery.trim().toLowerCase()))&&<p className="panel-empty">没有匹配的空间，请换一个关键词。</p>}
       <div className="workspace-action-tabs" role="group" aria-label="工作空间操作"><button aria-pressed={spaceAction==='create'} onClick={()=>setSpaceAction('create')}><Plus size={15}/>创建空间</button><button aria-pressed={spaceAction==='join'} onClick={()=>setSpaceAction('join')}><Users size={15}/>加入空间</button></div>
       {spaceAction==='create'&&<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);run(async()=>{await cloudJson('/workspaces','POST',{name:data.get('name')});await load();await onRefresh();});}}>
        <label className="form-field">新工作空间名称<input name="name" required maxLength={160} disabled={busy} placeholder="例如：产品设计团队"/></label>
        <button className="primary" disabled={busy}><Plus size={15}/>创建团队空间</button>
       </form>}
       {spaceAction==='join'&&<form onSubmit={event=>{event.preventDefault();const data=new FormData(event.currentTarget);run(async()=>{await cloudJson('/accept','POST',{token:data.get('token')});await load();await onRefresh();});}}>
        <label className="form-field">邀请口令<input name="token" required maxLength={128} disabled={busy} placeholder="粘贴邀请口令"/></label>
        <button disabled={busy}><Users size={15}/>接受邀请</button>
       </form>}
      </div>}
    </div>}
   </div>
  </div>
 </Modal>;
}
