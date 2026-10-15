import {GitHubReleaseCheck} from './github-release-check';
import {useEffect,useState} from 'react';import type {DesktopUpdateState} from '../../../common/nbstore/src/updates';import {Modal} from './dialogs';import {ClientReleaseCheck} from './client-release-check';import {releaseTarget} from '../../../common/nbstore/src/releases';import {ChevronDown} from 'lucide-react';
export function useDesktopUpdates(){
 const [state,setState]=useState<DesktopUpdateState>(),[open,setOpen]=useState(false);
 useEffect(()=>{const bridge=window.mydocDesktop;if(!bridge)return;let alive=true,pending=false;const show=()=>{pending=true;resume();};const resume=()=>{if(pending&&!document.querySelector('dialog[open]')){pending=false;setOpen(true);}};
  void bridge.updateStatus().then(value=>{if(alive)setState(value);}).catch(()=>{});const unsubscribe=bridge.onUpdate(value=>setState(value));const menu=bridge.onShowUpdates(show);const observer=new MutationObserver(resume);observer.observe(document.body,{childList:true,subtree:true,attributes:true,attributeFilter:['open']});return()=>{alive=false;unsubscribe();menu();observer.disconnect();};
 },[]);
 return {state,open,setOpen,setState};
}
function VersionSummary({state}:{state:DesktopUpdateState}){return <dl className="update-summary"><div><dt>当前版本</dt><dd>MyDoc {state.version}</dd></div><div><dt>更新渠道</dt><dd>{state.channel}</dd></div>{state.latestVersion&&<div><dt>可用版本</dt><dd>{state.latestVersion}</dd></div>}</dl>;}
const labels:Record<DesktopUpdateState['phase'],string>={disabled:'更新渠道未配置',idle:'尚未检查更新',checking:'正在检查更新…','up-to-date':'当前已是最新版本',available:'发现新版本',downloading:'正在下载更新…',downloaded:'更新已下载，准备安装',preparing:'正在保存文档并生成保护备份…',installing:'正在安装更新，请稍候…',error:'更新未完成'};
function UpdateStatus({state}:{state:DesktopUpdateState}){return <div className="update-status"><strong role="status">{labels[state.phase]}</strong>{state.phase==='disabled'&&<p>此渠道尚未配置，可使用上方的 GitHub 或服务器检查下载更新。</p>}{state.message&&state.phase!=='disabled'&&<p role="alert" className="form-error">{state.message}</p>}{state.checkedAt&&<p>上次检查：{new Date(state.checkedAt).toLocaleString()}</p>}{state.phase==='downloading'&&<><progress max={100} value={state.percent??0} aria-label="更新下载进度"/><p>{Math.round(state.percent??0)}%{state.total?` · ${(state.transferred??0)/1024/1024<1?'<1':((state.transferred??0)/1024/1024).toFixed(1)} / ${(state.total/1024/1024).toFixed(1)} MB`:''}</p></>}</div>;}
export function DesktopUpdateDialog({state,onState,onClose}:{state?:DesktopUpdateState;onState:(state:DesktopUpdateState)=>void;onClose:()=>void}){
 const [requesting,setRequesting]=useState(false),[error,setError]=useState('');const busy=requesting||!!state&&['checking','downloading','preparing','installing'].includes(state.phase),locking=state?.phase==='preparing'||state?.phase==='installing';
 const readStatus=async()=>{setRequesting(true);setError('');try{onState(await window.mydocDesktop!.updateStatus());}catch{setError('无法读取版本信息，请重试。');}finally{setRequesting(false);}};
 useEffect(()=>{if(!state)void readStatus();},[]);
 const run=async(action:'check'|'download'|'install')=>{if(requesting)return;setRequesting(true);setError('');try{onState(await window.mydocDesktop!.updateAction(action));}catch{setError('更新请求失败，请稍后重试。');}finally{setRequesting(false);}};
 return <Modal title="版本与更新" className="version-update-dialog" onClose={()=>{if(!locking)onClose();}}>
  <div className="version-update-layout">
   <header className="client-version-header"><div><span>当前安装版本</span><strong>{state?'MyDoc '+state.version:'正在读取…'}</strong></div><small>{window.mydocDesktop?.arch?releaseTarget(window.mydocDesktop.platform,window.mydocDesktop.arch):'读取设备信息…'}</small></header>
   {error&&<p role="alert" className="release-feedback error">{error}</p>}
   {!state&&error&&<button className="release-secondary" disabled={requesting} onClick={()=>void readStatus()}>重试读取</button>}
   <GitHubReleaseCheck state={state} onState={onState} locked={busy||state?.phase==='downloaded'&&state.source!=='github'}/>
   <ClientReleaseCheck state={state} onState={onState} locked={busy}/>
   {state&&state.source==='configured'&&<details className="automatic-update-section" open={state.enabled?true:undefined}>
    <summary><span>自动更新渠道</span><small>{state.enabled?state.channel:'尚未启用'}</small><ChevronDown size={14}/></summary>
    <div className="desktop-update"><VersionSummary state={state}/><UpdateStatus state={state}/>
     <label className="update-auto"><input type="checkbox" disabled={busy||!state.enabled} checked={state.automaticChecks} onChange={event=>{const value=event.target.checked;setRequesting(true);setError('');void window.mydocDesktop!.setAutomaticUpdateChecks(value).then(onState).catch(()=>setError('设置保存失败，请重试。')).finally(()=>setRequesting(false));}}/>自动检查更新</label>
     <p className="organization-help">启用后，启动及每 6 小时检查。安装前保存文档并生成数据库保护备份。</p>
     <div className="update-actions"><button className="release-secondary" disabled={busy||!state.enabled||state.phase==='downloaded'} onClick={()=>void run('check')}>{state.retry==='check'?'重试检查':'检查更新'}</button>{(state.phase==='available'||state.phase==='error'&&state.retry==='download')&&<button className="primary" disabled={busy} onClick={()=>void run('download')}>{state.retry==='download'?'重试下载':'下载更新'}</button>}{state.phase==='downloaded'&&<button className="primary" disabled={busy} onClick={()=>void run('install')}>{state.retry==='install'?'重试安装':'重启并安装'}</button>}</div>
    </div>
   </details>}
  </div>
  <div className="dialog-actions"><button disabled={locking} onClick={onClose}>{state?.phase==='downloaded'?'稍后安装':'关闭'}</button></div>
 </Modal>;
}
