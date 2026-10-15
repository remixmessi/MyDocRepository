import {useEffect,useId,useState} from 'react';
import {Download,RefreshCw,Server} from 'lucide-react';
import {projectPreferences} from '@mydoc/nbstore';
import type {DesktopUpdateState} from '../../../common/nbstore/src/updates';
import {normalizeUpdateServer} from '../../../common/nbstore/src/server-updates';
import {releaseTarget} from '../../../common/nbstore/src/releases';
import {formatBytes} from './cloud-settings';
import './release-styles.css';

const labels:Record<DesktopUpdateState['phase'],string>={disabled:'应用内更新不可用',idle:'尚未检查服务器更新',checking:'正在检查服务器版本…','up-to-date':'当前已是最新版本',available:'发现新版本',downloading:'正在下载更新…',downloaded:'更新已下载，可以重启安装',preparing:'正在保存文档并生成保护备份…',installing:'正在安装更新，请稍候…',error:'服务器更新未完成'};
export function ClientReleaseCheck({state,onState,locked}:{state?:DesktopUpdateState;onState:(state:DesktopUpdateState)=>void;locked:boolean}){
 const initial=()=>state?.source==='server'?state.serverUrl??'':projectPreferences.getItem('update-server-url')??'';
 const [server,setServer]=useState(initial),[draft,setDraft]=useState(initial),[requesting,setRequesting]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState('');
 const inputId=useId(),active=state?.source==='server'&&state.serverUrl===server,phase=active?state.phase:'idle';
 const busy=requesting||locked,retained=state?.phase==='downloaded',changed=draft.trim()!==server;
 useEffect(()=>{if(state?.source==='server'&&state.serverUrl&&draft===server){setServer(state.serverUrl);setDraft(state.serverUrl);}},[state?.serverUrl,state?.source]);
 const platform=window.mydocDesktop?.platform,arch=window.mydocDesktop?.arch,target=platform&&arch?releaseTarget(platform,arch):'正在读取客户端架构';
 const save=async()=>{
  setRequesting(true);setError('');setNotice('');try{
   const next=draft.trim()?normalizeUpdateServer(draft):'';await projectPreferences.setItem('update-server-url',next);setServer(next);setDraft(next);setNotice(next?'更新服务器地址已保存':'更新服务器地址已清除');
  }catch(error){setError(error instanceof Error?error.message:'地址保存失败');}finally{setRequesting(false);}
 };
 const run=async(action:'check'|'download'|'install')=>{
  if(busy)return;setRequesting(true);setError('');setNotice('');
  try{onState(action==='check'?await window.mydocDesktop!.checkServerUpdates(server):await window.mydocDesktop!.updateAction(action));}
  catch{setError('服务器更新请求失败，请稍后重试。');}finally{setRequesting(false);}
 };
 return <section className="client-release-check" aria-label="服务器版本检查" aria-busy={busy}>
  <div className="release-section-head"><div><h3><Server size={17}/>服务器更新</h3><p>从指定服务器下载更新，在客户端内重启安装。</p></div><span className="release-badge">应用内更新</span></div>
  <div className="client-target"><span>{target}</span><small>自动匹配系统与架构</small></div>
  <form onSubmit={event=>{event.preventDefault();void save();}}><label htmlFor={inputId} className="client-server-label">更新服务器地址</label><div className="client-server-input"><input id={inputId} aria-label="更新服务器地址" type="text" inputMode="url" value={draft} placeholder="192.168.1.20:8080 或 docs.example.com" disabled={busy||retained} onChange={event=>setDraft(event.target.value)}/><button className="release-secondary" disabled={busy||retained||!changed}>保存地址</button></div><small className="release-muted">支持 IP、域名与完整 HTTP / HTTPS 地址。</small></form>
  {notice&&<p className="release-feedback" role="status">{notice}</p>}{error&&<p className="release-feedback error" role="alert">{error}</p>}
  <div className="client-check-actions"><button className="primary" disabled={busy||retained||!server||changed||!state} onClick={()=>void run('check')}><RefreshCw size={15}/>{phase==='checking'?'正在检查服务器版本…':active&&state.retry==='check'?'重新检查':'检查服务器版本'}</button><small>{changed?'先保存地址，再检查版本':server?'已配置服务器':'填写并保存地址后即可检查'}</small></div>
  {active&&<>
   <div className="client-release-result"><strong role="status">{labels[phase]}</strong>{state.latestVersion&&<h4>{state.latestVersion}</h4>}
    {state.message&&<p className={phase==='disabled'?'release-muted':'release-feedback error'} role={phase==='disabled'?undefined:'alert'}>{state.message}</p>}
    {state.checkedAt&&<small className="release-muted">上次检查：{new Date(state.checkedAt).toLocaleString('zh-CN')}</small>}
    {state.releaseNotes&&<details className="release-details"><summary>更新说明</summary><div><p>{state.releaseNotes}</p></div></details>}
    {phase==='downloading'&&<><progress max={100} value={state.percent??0} aria-label="服务器更新下载进度"/><p>{Math.round(state.percent??0)}% · {formatBytes(state.transferred??0)}{state.total?' / '+formatBytes(state.total):''}</p></>}
    {(phase==='available'||phase==='error'&&state.retry==='download')&&<button className="primary" disabled={busy||changed} onClick={()=>void run('download')}><Download size={15}/>{state.retry==='download'?'重试下载更新':'下载并更新'}</button>}
    {phase==='downloaded'&&<button className="primary" disabled={busy} onClick={()=>void run('install')}>{state.retry==='install'?'重试安装':'重启并安装'}</button>}
   </div>
   <label className="update-auto"><input type="checkbox" disabled={busy||!state.enabled} checked={state.automaticChecks} onChange={event=>{setRequesting(true);setError('');void window.mydocDesktop!.setAutomaticUpdateChecks(event.target.checked).then(onState).catch(()=>setError('更新设置保存失败，请重试。')).finally(()=>setRequesting(false));}}/>自动检查服务器更新</label>
   <p className="release-muted">启动后及每 6 小时检查所选来源。安装前自动保存文档并生成数据库保护备份。macOS 更新需由管理员上传 ZIP，Windows 需上传 EXE。</p>
  </>}
 </section>;
}
