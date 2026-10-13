import {useCallback,useEffect,useState} from 'react';
import {ChevronDown,Palette,UserRound,Users,Package} from 'lucide-react';
import {cloudJson,type CloudAccount,type CloudMailStatus} from '../../../common/nbstore/src/cloud';
import {Modal} from './dialogs';
import {ServerConnections} from './server-connections';
import {LocalAccountPanel} from './local-account-ui';
import {CloudLogin} from './cloud-ui';
import {AccountSettings} from './account-settings';
import {ServerAdminSettings} from './server-admin-settings';
import {ReleaseManagement} from './release-management';
import type {ThemeMode} from './theme';

type Tab='local'|'server'|'appearance'|'users'|'releases';
export function SettingsDialog({onClose,onRefresh,theme,onThemeSave}:{onClose:()=>void;onRefresh:()=>Promise<void>;theme:ThemeMode;onThemeSave:(mode:ThemeMode)=>Promise<void>}){
 const [tab,setTab]=useState<Tab>('local'),[expanded,setExpanded]=useState(true),[user,setUser]=useState<CloudAccount|null>(),[mail,setMail]=useState<CloudMailStatus>({configured:false}),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const load=useCallback(async()=>{const result=await cloudJson<{user:CloudAccount|null;mail:CloudMailStatus}>('/session');setUser(result.user);setMail(result.mail??{configured:false});},[]);
 useEffect(()=>{let alive=true;void cloudJson<{user:CloudAccount|null;mail:CloudMailStatus}>('/session').then(result=>{if(alive){setUser(result.user);setMail(result.mail??{configured:false});}}).catch(()=>{if(alive)setUser(null);});return()=>{alive=false;};},[]);
 const run=(work:()=>Promise<void>)=>{setBusy(true);setError('');void work().catch(e=>setError(e instanceof Error?e.message:'操作失败，请重试')).finally(()=>setBusy(false));};
 const select=(value:Tab)=>{if(!busy){setTab(value);setError('');}};
 return <Modal title="设置" className="workspace-settings-dialog unified-settings" onClose={()=>{if(!busy)onClose();}}><div className="settings-layout">
  <nav className="settings-navigation" aria-label="设置导航">
   <p className="settings-nav-label">个人设置</p>
   <button className="settings-nav-group" aria-expanded={expanded} aria-controls="settings-account-options" onClick={()=>setExpanded(!expanded)}><UserRound size={17}/><span>账号</span><ChevronDown size={15} className={expanded?'expanded':''}/></button>
   <div id="settings-account-options" hidden={!expanded} className="settings-nav-children"><button aria-current={tab==='local'?'page':undefined} disabled={busy} onClick={()=>select('local')}>本地账号</button><button aria-current={tab==='server'?'page':undefined} disabled={busy} onClick={()=>select('server')}>服务器账号</button></div>
   <button aria-current={tab==='appearance'?'page':undefined} disabled={busy} onClick={()=>select('appearance')}><Palette size={17}/><span>外观</span></button>
   <p className="settings-nav-label">管理</p>
   <button aria-current={tab==='users'?'page':undefined} disabled={busy} onClick={()=>select('users')}><Users size={17}/><span>用户管理</span></button>
   <button aria-current={tab==='releases'?'page':undefined} disabled={busy} onClick={()=>select('releases')}><Package size={17}/><span>版本维护</span></button>
   <p className="settings-nav-note">团队空间的成员与邀请，请在工作区菜单中管理。</p>
  </nav>
  <div className="settings-body" aria-busy={busy}>
   {error&&<p className="form-error" role="alert">{error}</p>}
   {tab==='local'&&<LocalAccountPanel/>}
   {tab==='appearance'&&<AppearancePanel mode={theme} save={onThemeSave}/>}
   {tab==='server'&&<>
    <h2>服务器账号</h2><p className="organization-help">连接服务器，登录并管理账号资料。</p>
    <ServerConnections busy={busy} prepare={onRefresh}/>
    {user===undefined?<p role="status">正在读取服务器账号…</p>:user?<><AccountSettings key={user.id} account={user} mail={mail} busy={busy} run={run} onBack={()=>select('local')} onChanged={load} onSignedOut={()=>location.reload()}/><button disabled={busy} onClick={()=>run(async()=>{await onRefresh();await cloudJson('/logout','POST');location.reload();})}>退出服务器账号</button></>:<CloudLogin beforeSession={onRefresh} onLogin={()=>location.reload()}/>}
   </>}
   {(tab==='users'||tab==='releases')&&(user===undefined?<p role="status">正在读取管理权限…</p>:!user?.serverAdmin?<section><h2>{tab==='users'?'用户管理':'版本维护'}</h2><p className="organization-help">登录服务器管理员账号后可使用此功能。</p><button onClick={()=>{setExpanded(true);select('server');}}>登录服务器账号</button></section>:tab==='users'?<ServerAdminSettings accountId={user.id} busy={busy} run={run} onBack={()=>select('server')}/>:<ReleaseManagement accountId={user.id} busy={busy} run={run} onBack={()=>select('server')}/>)}
  </div>
 </div></Modal>;
}

function AppearancePanel({mode,save}:{mode:ThemeMode;save:(mode:ThemeMode)=>Promise<void>}){
 const [busy,setBusy]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState('');
 return <section><h2>外观设置</h2><p className="organization-help">设置当前应用的配色，保存后立即生效。</p><div className="settings-preference-row"><div><h3>配色方案</h3><p>跟随系统，或固定使用浅色与深色。</p></div><div className="settings-theme-options" role="group" aria-label="配色方案">{([{id:'system',label:'跟随系统'},{id:'light',label:'浅色'},{id:'dark',label:'深色'}] as const).map(option=><button key={option.id} aria-pressed={mode===option.id} disabled={busy} onClick={()=>{setBusy(true);setError('');setNotice('');void save(option.id).then(()=>setNotice('外观已保存')).catch(()=>setError('保存失败，请重试')).finally(()=>setBusy(false));}}>{option.label}</button>)}</div></div>{error&&<p role="alert" className="form-error">{error}</p>}{notice&&<p role="status" className="form-notice">{notice}</p>}</section>;
}
