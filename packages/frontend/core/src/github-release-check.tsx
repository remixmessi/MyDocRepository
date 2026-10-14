import {useEffect,useRef,useState} from 'react';
import {Download,ExternalLink,Github,RefreshCw} from 'lucide-react';
import {checkGitHubRelease,githubReleasesPage,type GitHubRelease} from '../../../common/nbstore/src/github-releases';
import {compareReleaseVersions,releaseTarget,validReleaseVersion} from '../../../common/nbstore/src/releases';
import {formatBytes} from './cloud-settings';
import './release-styles.css';

export function GitHubReleaseCheck({current,locked}:{current?:string;locked:boolean}) {
 const [release,setRelease]=useState<GitHubRelease|null>(),[busy,setBusy]=useState(false),[error,setError]=useState(''),[checkedAt,setCheckedAt]=useState<number>();
 const request=useRef<AbortController>(),alive=useRef(true);
 useEffect(()=>{alive.current=true;return()=>{alive.current=false;request.current?.abort();};},[]);
 const platform=window.mydocDesktop?.platform,arch=window.mydocDesktop?.arch;
 const target=platform&&arch?releaseTarget(platform,arch):'正在读取客户端架构';
 const check=async()=>{
  if(request.current||locked)return;
  if(!platform||!['x64','arm64'].includes(arch??'')){setError('无法识别客户端架构，请重新打开应用。');return;}
  const controller=new AbortController();request.current=controller;setBusy(true);setError('');setRelease(undefined);
  const timer=setTimeout(()=>controller.abort(),15000);
  try {const result=await checkGitHubRelease(platform,arch!,controller.signal);if(alive.current){setRelease(result);setCheckedAt(Date.now());}}
  catch(error){if(alive.current)setError(controller.signal.aborted?'GitHub 连接超时，请检查网络后重试。':error instanceof TypeError?'无法连接 GitHub，请检查网络后重试。':error instanceof Error?error.message:'GitHub 版本检查失败，请重试。');}
  finally {clearTimeout(timer);request.current=undefined;if(alive.current)setBusy(false);}
 };
 const comparison=release&&current&&validReleaseVersion(current)?compareReleaseVersions(release.version,current):undefined;
 return <section className="client-release-check github-release-check" aria-label="GitHub 版本检查" aria-busy={busy}>
  <div className="release-section-head"><div><h3><Github size={17}/>GitHub 更新</h3><p>从 remixmessi/MyDocRepository 检查最新正式版本。</p></div><span className="release-badge">手动检查</span></div>
  <div className="client-target"><span>{target}</span><small>安装包按此系统与架构匹配</small></div>
  {error&&<p className="release-feedback error" role="alert">{error}</p>}
  <div className="client-check-actions"><button className="primary" disabled={busy||locked||!current} onClick={()=>void check()}><RefreshCw size={15}/>{busy?'正在检查 GitHub…':error?'重试 GitHub 检查':'检查 GitHub 版本'}</button><a className="release-secondary" href={githubReleasesPage} target="_blank" rel="noreferrer"><ExternalLink size={15}/>查看 Release 页面</a></div>
  {checkedAt&&<p className="release-muted">上次检查：{new Date(checkedAt).toLocaleString('zh-CN')}</p>}
  {release===null&&<div className="client-release-result" role="status"><strong>暂无公开正式版本</strong><p>尚未发布正式 Release，或仓库无法公开访问。可打开 Release 页面查看。</p></div>}
  {release&&<div className="client-release-result"><div><strong role="status">{comparison===undefined?'GitHub 可用版本':comparison>0?'发现新版本':comparison===0?'当前已是最新版本':'当前版本高于 GitHub 正式版本'}</strong><h4>{release.version}</h4></div>
   <div className="client-release-meta"><span>{target}</span>{release.publishedAt&&<span>{new Date(release.publishedAt).toLocaleDateString('zh-CN')}</span>}</div>
   {release.notes&&<details className="release-details"><summary>更新说明</summary><div><p>{release.notes}</p></div></details>}
   {release.installer?<><p className="github-installer-name">{release.installer.name} · {formatBytes(release.installer.bytes)}</p>{release.installer.sha256&&<details className="release-details"><summary>SHA-256 校验信息</summary><div><code>{release.installer.sha256}</code></div></details>}{!locked&&<a className="release-secondary" href={release.installer.url} target="_blank" rel="noreferrer"><Download size={15}/>下载 GitHub 安装包</a>}<small className="release-muted">在浏览器下载后手动安装。安装前请保存文档并退出客户端。</small></>:<p role="status">此版本尚未提供 {target} 的安装包，请查看 Release 页面。</p>}
  </div>}
 </section>;
}
