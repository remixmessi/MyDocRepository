import {useEffect,useId,useRef,useState} from 'react';
import {CheckCircle2,Download,RefreshCw,Server} from 'lucide-react';
import {projectPreferences} from '@mydoc/nbstore';
import {compareReleaseVersions,releaseTarget,validReleaseVersion,type DesktopRelease} from '../../../common/nbstore/src/releases';
import {formatBytes} from './cloud-settings';
import './release-styles.css';

export function ClientReleaseCheck({current,locked}:{current?:string;locked:boolean}){
 const [server,setServer]=useState(()=>projectPreferences.getItem('update-server-url')??''),[draft,setDraft]=useState(()=>projectPreferences.getItem('update-server-url')??''),[release,setRelease]=useState<DesktopRelease|null>(),[busy,setBusy]=useState(false),[error,setError]=useState(''),[notice,setNotice]=useState(''),[checkedAt,setCheckedAt]=useState<number>();
 const request=useRef<AbortController>(),alive=useRef(true);
 const inputId=useId();
 useEffect(()=>{alive.current=true;return()=>{alive.current=false;request.current?.abort();};},[]);
 const platform=window.mydocDesktop?.platform,arch=window.mydocDesktop?.arch;
 const target=platform&&arch?releaseTarget(platform,arch):'正在读取客户端架构';
 const changed=draft.trim()!==server;
 const check=async()=>{
  if(!platform||!['x64','arm64'].includes(arch??'')){setError('无法识别客户端架构，请重新打开应用。');return;}
  setBusy(true);setError('');setNotice('');setRelease(undefined);const controller=new AbortController();request.current=controller;
  const timer=setTimeout(()=>controller.abort(new Error('连接超时，请检查服务器地址和网络')),15000);
  try{
   const response=await fetch(server+'/api/cloud/releases/latest?'+new URLSearchParams({platform,arch:arch!}),{headers:{Accept:'application/json'},signal:controller.signal});
   if(!response.ok)throw new Error('更新服务器返回错误（'+response.status+'）');const value=await response.json() as DesktopRelease|null;
   if(value&&(!validReleaseVersion(value.version)||value.platform!==platform||value.arch!==arch&&!(platform!=='darwin'&&value.arch==='')||!/^\/api\/cloud\/releases\/download\?id=[a-f0-9-]+$/.test(value.downloadPath)))throw new Error('服务器返回的安装包与当前系统或架构不匹配');
   if(alive.current){setRelease(value);setCheckedAt(Date.now());}
  }catch(error){if(alive.current)setError(error instanceof Error?error.message:'无法连接更新服务器');}finally{clearTimeout(timer);if(alive.current)setBusy(false);}
 };
 const save=async()=>{
  setBusy(true);setError('');setNotice('');try{
   const input=draft.trim();let next='';if(input){const value=new URL(/^[a-z][a-z\d+.-]*:\/\//i.test(input)?input:'http://'+input);if(!['http:','https:'].includes(value.protocol)||value.username||value.password||value.search||value.hash)throw new Error('请填写 IP、域名或 http／https 地址，不含用户名、查询参数或锚点');next=value.origin+value.pathname.replace(/\/+$/,'');}
   await projectPreferences.setItem('update-server-url',next);if(alive.current){setServer(next);setDraft(next);setRelease(undefined);setCheckedAt(undefined);setNotice(next?'更新服务器地址已保存':'更新服务器地址已清除');}
  }catch(error){if(alive.current)setError(error instanceof Error?error.message:'地址保存失败');}finally{if(alive.current)setBusy(false);}
 };
 const newer=release&&current&&validReleaseVersion(current)&&compareReleaseVersions(release.version,current)>0;
 const same=release&&current&&validReleaseVersion(current)&&compareReleaseVersions(release.version,current)===0;
 return <section className="client-release-check" aria-label="服务器版本检查">
  <div className="release-section-head"><div><h3><Server size={17}/>服务器更新</h3><p>连接 Web 服务器，检查适用于这台设备的安装包。</p></div><span className="release-badge">手动检查</span></div>
  <div className="client-target"><span>{target}</span><small>安装包按此系统与架构匹配</small></div>
  <form onSubmit={event=>{event.preventDefault();void save();}}><label htmlFor={inputId} className="client-server-label">更新服务器地址</label><div className="client-server-input"><input id={inputId} aria-label="更新服务器地址" type="text" inputMode="url" value={draft} placeholder="192.168.1.20:8080 或 docs.example.com" disabled={busy||locked} onChange={event=>setDraft(event.target.value)}/><button className="release-secondary" disabled={busy||locked||!changed}>保存地址</button></div><small className="release-muted">支持 IP、域名与完整 HTTP / HTTPS 地址。</small></form>
  {notice&&<p className="release-feedback" role="status">{notice}</p>}{error&&<p className="release-feedback error" role="alert">{error}</p>}
  <div className="client-check-actions"><button className="primary" disabled={busy||locked||!server||changed||!current} onClick={()=>void check()}><RefreshCw size={15}/>{busy?'正在处理…':error?'重新检查':'检查服务器版本'}</button><small>{changed?'先保存地址，再检查版本':checkedAt?'上次检查：'+new Date(checkedAt).toLocaleTimeString('zh-CN',{hour12:false}):server?'已配置服务器':'填写并保存地址后即可检查'}</small></div>
  {release===null&&<div className="client-release-result" role="status"><strong>暂无匹配的安装包</strong><p>服务器尚未发布 {target} 版本。</p></div>}
  {release&&<div className="client-release-result"><div className="release-section-head"><div><strong role="status">{newer?'发现新版本':same?'当前已是最新版本':'服务器可用版本'}</strong><h4>{release.version}</h4></div><CheckCircle2 size={22}/></div><div className="client-release-meta"><span>{releaseTarget(release.platform,release.arch)}</span><span>{formatBytes(release.bytes)} · {new Date(release.created).toLocaleDateString('zh-CN')}</span></div>{release.notes&&<p className="client-release-notes">{release.notes}</p>}<details className="release-details"><summary>文件与校验信息</summary><div><p>{release.filename}</p><code>{release.sha256}</code></div></details><a className="release-secondary" href={server+release.downloadPath} download><Download size={15}/>下载安装包</a><small className="release-muted">下载后手动安装，此方式不会自动重启应用。</small></div>}
 </section>;
}
