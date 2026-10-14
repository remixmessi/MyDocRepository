import {useCallback,useEffect,useState} from 'react';
import {ChevronDown,Palette,UserRound,Users,Package,Tags} from 'lucide-react';
import {cloudJson,cloudServerUrl,type CloudAccount,type CloudMailStatus,type CloudWorkspace} from '../../../common/nbstore/src/cloud';
import {Modal} from './dialogs';
import {ServerConnections} from './server-connections';
import {LocalAccountPanel} from './local-account-ui';
import {CloudLogin} from './cloud-ui';
import {AccountSettings} from './account-settings';
import {ServerAdminSettings} from './server-admin-settings';
import {WebReleaseSettings,DesktopVersionSettings} from './version-settings';
import type {ThemeMode} from './theme';
import type {LocalWorkspace} from './workspace';
import {OrganizationPanel} from './organization-ui';

export type SettingsTab='local'|'server'|'appearance'|'organization'|'users'|'releases';
export function SettingsDialog({onClose,onRefresh,onCloudSessionChanged,theme,onThemeSave,workspaces,activeWorkspaceId,workspace,initialTab='local'}:{onClose:()=>void;onRefresh:()=>Promise<void>;onCloudSessionChanged?:()=>Promise<void>;theme:ThemeMode;onThemeSave:(mode:ThemeMode)=>Promise<void>;workspaces:CloudWorkspace[];activeWorkspaceId?:string;workspace?:LocalWorkspace|null;initialTab?:SettingsTab}){
 const [tab,setTab]=useState<SettingsTab>(initialTab),[expanded,setExpanded]=useState(true),[user,setUser]=useState<CloudAccount|null>(),[mail,setMail]=useState<CloudMailStatus>({configured:false}),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const [sessionError,setSessionError]=useState('');
 const load=useCallback(async()=>{if(window.mydocDesktop&&!cloudServerUrl()){setSessionError('');setUser(null);setMail({configured:false});return;}setSessionError('');try{const result=await cloudJson<{user:CloudAccount|null;mail:CloudMailStatus}>('/session');setUser(result.user);setMail(result.mail??{configured:false});}catch(e){setSessionError(e instanceof Error?e.message:'服务器会话读取失败');}},[]);
 useEffect(()=>{if(window.mydocDesktop&&!cloudServerUrl()){setUser(null);setMail({configured:false});return;}let alive=true;void cloudJson<{user:CloudAccount|null;mail:CloudMailStatus}>('/session').then(result=>{if(alive){setUser(result.user);setMail(result.mail??{configured:false});}}).catch(e=>{if(alive)setSessionError(e instanceof Error?e.message:'服务器会话读取失败');});return()=>{alive=false;};},[]);
 const run=(work:()=>Promise<void>)=>{setBusy(true);setError('');void work().catch(e=>setError(e instanceof Error?e.message:'操作失败，请重试')).finally(()=>setBusy(false));};
 const select=(value:SettingsTab)=>{if(!busy){setTab(value);setError('');}};
 return <Modal title="设置" className="workspace-settings-dialog unified-settings" onClose={()=>{if(!busy)onClose();}}><div className="settings-layout">
  <nav className="settings-navigation" aria-label="设置导航">
   <p className="settings-nav-label">个人设置</p>
   <button className="settings-nav-group" aria-expanded={expanded} aria-controls="settings-account-options" onClick={()=>setExpanded(!expanded)}><UserRound size={17}/><span>账号</span><ChevronDown size={15} className={expanded?'expanded':''}/></button>
   <div id="settings-account-options" hidden={!expanded} className="settings-nav-children"><button aria-current={tab==='local'?'page':undefined} disabled={busy} onClick={()=>select('local')}>本地账号</button><button aria-current={tab==='server'?'page':undefined} disabled={busy} onClick={()=>select('server')}>服务器账号</button></div>
   <button aria-current={tab==='appearance'?'page':undefined} disabled={busy} onClick={()=>select('appearance')}><Palette size={17}/><span>外观</span></button>
   <p className="settings-nav-label">管理</p>
   <button aria-current={tab==='organization'?'page':undefined} disabled={busy||!workspace||workspace.readOnly} onClick={()=>select('organization')}><Tags size={17}/><span>标签与集合</span></button>
   <button aria-current={tab==='users'?'page':undefined} disabled={busy} onClick={()=>select('users')}><Users size={17}/><span>用户管理</span></button>
   <button aria-current={tab==='releases'?'page':undefined} disabled={busy} onClick={()=>select('releases')}><Package size={17}/><span>{window.mydocDesktop?'版本与更新':'版本维护'}</span></button>
   <p className="settings-nav-note">团队空间的成员与邀请，请在工作区菜单中管理。</p>
  </nav>
  <div className="settings-body" aria-busy={busy}>
   {error&&<p className="form-error" role="alert">{error}</p>}
   {tab==='local'&&<LocalAccountPanel/>}
   {tab==='appearance'&&<AppearancePanel mode={theme} save={onThemeSave}/>}
   {tab==='organization'&&<section><h2>标签与集合</h2><p className="organization-help">维护当前工作区的标签和集合。文档筛选仍在侧栏中进行。</p>{!workspace?<p role="status">正在打开当前工作区…</p>:workspace.readOnly?<p role="status">当前工作区为只读，无法维护标签与集合。</p>:<OrganizationPanel key={workspace.id} workspace={workspace} onBusyChange={setBusy}/>}</section>}
   {tab==='server'&&<>
    <h2>服务器账号</h2><p className="organization-help">连接服务器，登录并管理账号资料。</p>
    <ServerConnections busy={busy} prepare={onRefresh} onServerChange={load}/>
    {window.mydocDesktop&&!cloudServerUrl()?<p role="status">请先添加远程服务器，再登录服务器账号。</p>:sessionError?<p role="alert">{sessionError}<button disabled={busy} onClick={()=>run(load)}>重试读取账号</button></p>:user===undefined?<p role="status">正在读取服务器账号…</p>:user?<><AccountSettings key={user.id} account={user} mail={mail} busy={busy} run={run} onBack={()=>select('local')} onChanged={load} onSignedOut={async()=>{await onCloudSessionChanged?.();await load();onClose();}}/><button disabled={busy} onClick={()=>run(async()=>{await onRefresh();await cloudJson('/logout','POST');await onCloudSessionChanged?.();await load();onClose();})}>退出服务器账号</button></>:<CloudLogin beforeSession={onRefresh} onLogin={async()=>{await load();await onCloudSessionChanged?.();}}/>}
   </>}
   {tab==='users'&&<><p className="organization-help">当前团队服务器：{cloudServerUrl()||(window.mydocDesktop?'未连接远程服务器':location.origin)}</p>{sessionError?<p role="alert">账号读取失败：{sessionError}<button disabled={busy} onClick={()=>run(load)}>重试读取账号</button></p>:user===undefined?<p role="status">正在读取管理权限…</p>:!user?<section><h2>用户管理</h2><p className="organization-help">服务器地址已选定，请登录该服务器账号。</p><button onClick={()=>{setExpanded(true);select('server');}}>登录服务器账号</button></section>:!user.serverAdmin?<section><h2>用户管理</h2><p role="status">已登录 {user.email}，当前账号没有服务器用户管理权限。请联系该服务器管理员授权。</p></section>:<ServerAdminSettings accountId={user.id} busy={busy} run={run} workspaces={workspaces} activeWorkspaceId={activeWorkspaceId} onBack={()=>select('server')}/>}</>}
   {tab==='releases'&&(window.mydocDesktop?<DesktopVersionSettings busy={busy}/>:<WebReleaseSettings busy={busy} run={run}/>)}
  </div>
 </div></Modal>;
}

function AppearancePanel({mode,save}:{mode:ThemeMode;save:(mode:ThemeMode)=>Promise<void>}){
 const [busy,setBusy]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState('');
 return <section><h2>外观设置</h2><p className="organization-help">设置当前应用的配色，保存后立即生效。</p><div className="settings-preference-row"><div><h3>配色方案</h3><p>跟随系统，或固定使用浅色与深色。</p></div><div className="settings-theme-options" role="group" aria-label="配色方案">{([{id:'system',label:'跟随系统'},{id:'light',label:'浅色'},{id:'dark',label:'深色'}] as const).map(option=><button key={option.id} aria-pressed={mode===option.id} disabled={busy} onClick={()=>{setBusy(true);setError('');setNotice('');void save(option.id).then(()=>setNotice('外观已保存')).catch(()=>setError('保存失败，请重试')).finally(()=>setBusy(false));}}>{option.label}</button>)}</div></div>{error&&<p role="alert" className="form-error">{error}</p>}{notice&&<p role="status" className="form-notice">{notice}</p>}</section>;
}
