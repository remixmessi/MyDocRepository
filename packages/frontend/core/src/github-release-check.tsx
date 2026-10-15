import {useState} from 'react';
import {Download,ExternalLink,Github,RefreshCw} from 'lucide-react';
import {githubReleasesPage} from '../../../common/nbstore/src/github-releases';
import type {DesktopUpdateState} from '../../../common/nbstore/src/updates';
import {releaseTarget} from '../../../common/nbstore/src/releases';
import {formatBytes} from './cloud-settings';
import './release-styles.css';

const statusLabels:Record<DesktopUpdateState['phase'],string>={disabled:'应用内更新不可用',idle:'尚未检查 GitHub 更新',checking:'正在检查 GitHub…','up-to-date':'当前已是最新版本',available:'发现新版本',downloading:'正在下载更新…',downloaded:'更新已下载，可以重启安装',preparing:'正在保存文档并生成保护备份…',installing:'正在安装更新，请稍候…',error:'GitHub 更新未完成'};
export function GitHubReleaseCheck({state,onState,locked}:{state?:DesktopUpdateState;onState:(state:DesktopUpdateState)=>void;locked:boolean}) {
 const [requesting,setRequesting]=useState(false),[error,setError]=useState('');
 const active=state?.source==='github',phase=active?state.phase:'idle';
 const platform=window.mydocDesktop?.platform,arch=window.mydocDesktop?.arch;
 const target=platform&&arch?releaseTarget(platform,arch):'正在读取客户端架构';
 const busy=requesting||locked,enabled=!!state&&(active?state.enabled:['darwin','win32'].includes(platform??''));
 const run=async(action:'check-github'|'download'|'install')=>{
  if(busy)return;setRequesting(true);setError('');
  try{onState(await window.mydocDesktop!.updateAction(action));}
  catch{setError('更新请求失败，请稍后重试。');}finally{setRequesting(false);}
 };
 return <section className="client-release-check github-release-check" aria-label="GitHub 版本检查" aria-busy={busy}>
  <div className="release-section-head"><div><h3><Github size={17}/>GitHub 更新</h3><p>从 remixmessi/MyDocRepository 下载更新，在客户端内重启安装。</p></div><span className="release-badge">应用内更新</span></div>
  <div className="client-target"><span>{target}</span><small>自动匹配系统与架构</small></div>
  <div className="client-check-actions"><button className="primary" disabled={busy||!enabled||phase==='downloaded'} onClick={()=>void run('check-github')}><RefreshCw size={15}/>{phase==='checking'?'正在检查 GitHub…':active&&state.retry==='check'?'重试 GitHub 检查':'检查 GitHub 版本'}</button><a className="release-secondary" href={githubReleasesPage} target="_blank" rel="noreferrer"><ExternalLink size={15}/>查看 Release 页面</a></div>
  {error&&<p className="release-feedback error" role="alert">{error}</p>}
  {active&&<>
   <div className="client-release-result"><strong role="status">{statusLabels[phase]}</strong>
    {state.latestVersion&&<h4>{state.latestVersion}</h4>}
    {state.message&&<p className={phase==='disabled'?'release-muted':'release-feedback error'} role={phase==='disabled'?undefined:'alert'}>{state.message}</p>}
    {state.checkedAt&&<small className="release-muted">上次检查：{new Date(state.checkedAt).toLocaleString('zh-CN')}</small>}
    {state.releaseNotes&&<details className="release-details"><summary>更新说明</summary><div><p>{state.releaseNotes}</p></div></details>}
    {phase==='downloading'&&<><progress max={100} value={state.percent??0} aria-label="GitHub 更新下载进度"/><p>{Math.round(state.percent??0)}% · {formatBytes(state.transferred??0)}{state.total?' / '+formatBytes(state.total):''}</p></>}
    {(phase==='available'||phase==='error'&&state.retry==='download')&&<button className="primary" disabled={busy} onClick={()=>void run('download')}><Download size={15}/>{state.retry==='download'?'重试下载更新':'下载并更新'}</button>}
    {phase==='downloaded'&&<button className="primary" disabled={busy} onClick={()=>void run('install')}>{state.retry==='install'?'重试安装':'重启并安装'}</button>}
   </div>
   <label className="update-auto"><input type="checkbox" disabled={busy||!state.enabled} checked={state.automaticChecks} onChange={event=>{setRequesting(true);setError('');void window.mydocDesktop!.setAutomaticUpdateChecks(event.target.checked).then(onState).catch(()=>setError('更新设置保存失败，请重试。')).finally(()=>setRequesting(false));}}/>自动检查 GitHub 更新</label>
   <p className="release-muted">启动后及每 6 小时检查。安装前自动保存文档并生成数据库保护备份；下载完成后可稍后安装。</p>
  </>}
 </section>;
}
