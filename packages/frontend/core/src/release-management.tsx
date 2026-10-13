import {useCallback,useEffect,useId,useRef,useState} from 'react';
import {ArrowLeft,Check,ChevronDown,Download,FileUp,Package,Plus,RefreshCw,Trash2} from 'lucide-react';
import {cloudJson,cloudRequest,cloudUrl} from '../../../common/nbstore/src/cloud';
import {compareReleaseVersions,releasePlatforms,releaseTarget,releaseVersionPattern,validReleaseVersion,type DesktopRelease,type ReleaseArchitecture,type ReleasePlatform} from '../../../common/nbstore/src/releases';
import {Modal} from './dialogs';
import {formatBytes} from './cloud-settings';
import './release-styles.css';

const date=(value:number)=>new Date(value).toLocaleString('zh-CN',{hour12:false});
const downloadUrl=(item:DesktopRelease)=>cloudUrl(item.downloadPath.replace(/^\/api\/cloud/,''));
function latestFor(items:DesktopRelease[],platform:ReleasePlatform,arch:ReleaseArchitecture){
 return items.filter(item=>item.platform===platform&&item.arch===arch&&validReleaseVersion(item.version)).sort((a,b)=>compareReleaseVersions(b.version,a.version)||b.created-a.created)[0];
}

function ReleaseComposer({busy,onUpload,onCancel}:{busy:boolean;onUpload:(file:File,version:string,platform:ReleasePlatform,arch:ReleaseArchitecture,notes:string)=>void;onCancel:()=>void}){
 const [version,setVersion]=useState(''),[platform,setPlatform]=useState<ReleasePlatform>('darwin'),[arch,setArch]=useState<ReleaseArchitecture>('arm64'),[notes,setNotes]=useState(''),[file,setFile]=useState<File>(),[error,setError]=useState('');
 const input=useRef<HTMLInputElement>(null);
 const fileId=useId();
 return <form className="release-composer" aria-label="发布新版本" onSubmit={event=>{
  event.preventDefault();if(!validReleaseVersion(version.trim())){setError('请输入有效版本号，例如 1.2.3 或 1.2.3-beta.1');return;}if(!file||!file.size){setError('请选择非空安装包');return;}if(file.size>512*1024*1024){setError('安装包不能超过 512 MB');return;}setError('');onUpload(file,version.trim(),platform,arch,notes);
 }}>
  <div className="release-section-head"><div><h4>发布新版本</h4><p>每个系统与架构分别上传一个安装包。</p></div><span className="release-badge">{releaseTarget(platform,arch)}</span></div>
  <div className="release-fields">
   <label>版本号<input required autoFocus value={version} maxLength={100} pattern={releaseVersionPattern} placeholder="例如 1.2.3" disabled={busy} onChange={event=>setVersion(event.target.value)}/></label>
   <label>操作系统<select value={platform} disabled={busy} onChange={event=>setPlatform(event.target.value as ReleasePlatform)}>{Object.entries(releasePlatforms).map(([value,label])=><option key={value} value={value}>{label}</option>)}</select></label>
   <label>处理器架构<select value={arch} disabled={busy} onChange={event=>setArch(event.target.value as ReleaseArchitecture)}><option value="arm64">{platform==='darwin'?'Apple 芯片 · arm64':'arm64'}</option><option value="x64">{platform==='darwin'?'Intel · x64':'x64'}</option></select></label>
  </div>
  <label className="release-file-label" htmlFor={fileId}>安装包</label>
  <div className="release-file-picker"><FileUp size={22}/><div><strong>{file?.name??'选择安装包'}</strong><small>{file?formatBytes(file.size):'支持 DMG、ZIP、EXE 等安装文件，最大 512 MB'}</small></div><button type="button" className="release-secondary" disabled={busy} onClick={()=>input.current?.click()}>{file?'重新选择':'浏览文件'}</button><input id={fileId} ref={input} type="file" aria-label="安装包" disabled={busy} onChange={event=>{setFile(event.target.files?.[0]);setError('');}}/></div>
  <label className="release-notes-field">更新说明 <span>可选</span><textarea value={notes} onChange={event=>setNotes(event.target.value)} maxLength={4000} rows={3} disabled={busy} placeholder="介绍本次修复与改进，供客户端查看"/></label>
  {error&&<p className="release-feedback error" role="alert">{error}</p>}
  <div className="release-composer-actions"><small>发布后，匹配此架构的客户端即可检查到版本。</small><button type="button" disabled={busy} onClick={onCancel}>取消</button><button className="primary" disabled={busy||!file}><FileUp size={15}/>{busy?'正在上传…':'上传并发布'}</button></div>
 </form>;
}

function ReleaseRecord({item,latest,busy,onDelete}:{item:DesktopRelease;latest:boolean;busy:boolean;onDelete:()=>void}){
 return <li className="release-record">
  <div className="release-record-main"><div className="release-version"><strong>{item.version}</strong>{latest&&<span className="release-badge current">最新</span>}</div><span className={`release-target${!item.arch?' unspecified':''}`}>{releaseTarget(item.platform,item.arch)}</span><div className="release-package"><span title={item.filename}>{item.filename}</span><small>{formatBytes(item.bytes)} · {date(item.created)}</small></div><div className="release-record-actions"><a className="release-secondary" href={downloadUrl(item)} download aria-label={`下载 ${item.version} ${releaseTarget(item.platform,item.arch)}`}><Download size={15}/><span>下载</span></a><button className="release-icon-button" disabled={busy} onClick={onDelete} aria-label={`删除 ${item.version} ${releaseTarget(item.platform,item.arch)}`}><Trash2 size={16}/></button></div></div>
  <details className="release-details"><summary><ChevronDown size={14}/>版本详情</summary><div>{!item.arch&&item.platform==='darwin'&&<p className="release-legacy-note">此安装包未标注架构，不参与客户端自动匹配。请选择 x64 或 arm64 重新上传。</p>}{item.notes?<p>{item.notes}</p>:<p className="release-muted">未填写更新说明</p>}<label>SHA-256<code>{item.sha256}</code></label></div></details>
 </li>;
}

export function ReleaseManagement({accountId,busy,run,onBack}:{accountId:string;busy:boolean;run:(work:()=>Promise<void>)=>void;onBack:()=>void}){
 const [items,setItems]=useState<DesktopRelease[]>([]),[error,setError]=useState(''),[notice,setNotice]=useState(''),[loading,setLoading]=useState(true),[composing,setComposing]=useState(false),[filter,setFilter]=useState<ReleasePlatform|'all'>('all'),[archFilter,setArchFilter]=useState('all'),[pending,setPending]=useState<DesktopRelease>();
 const alive=useRef(true),generation=useRef(0);
 const load=useCallback(async()=>{const request=++generation.current;setLoading(true);setError('');try{const next=await cloudJson<DesktopRelease[]>('/admin/releases');if(alive.current&&request===generation.current)setItems(next);}catch(error){if(alive.current&&request===generation.current)setError(error instanceof Error?error.message:'无法读取版本列表');}finally{if(alive.current&&request===generation.current)setLoading(false);}},[]);
 useEffect(()=>{alive.current=true;void load();return()=>{alive.current=false;generation.current++;};},[accountId,load]);
 const latestIds=new Set(Object.keys(releasePlatforms).flatMap(platform=>(['arm64','x64'] as const).map(arch=>latestFor(items,platform as ReleasePlatform,arch)?.id)));
 const filtered=items.filter(item=>(filter==='all'||item.platform===filter)&&(archFilter==='all'||item.arch===archFilter));
 const upload=(file:File,version:string,platform:ReleasePlatform,arch:ReleaseArchitecture,notes:string)=>run(async()=>{
  setNotice('');const encoded=btoa(Array.from(new TextEncoder().encode(notes),byte=>String.fromCharCode(byte)).join(''));
  await cloudRequest('/admin/release',{method:'POST',headers:{'X-MyDoc-Version':version,'X-MyDoc-Platform':platform,'X-MyDoc-Arch':arch,'X-MyDoc-Filename':encodeURIComponent(file.name),'X-MyDoc-Notes':encoded},body:file});if(!alive.current)return;setComposing(false);setNotice(`${version} · ${releaseTarget(platform,arch)} 已发布`);await load();
 });
 return <section className="release-management" aria-label="版本维护">
  <button className="workspace-back" disabled={busy} onClick={onBack}><ArrowLeft size={15}/>返回设置</button>
  <header className="release-page-head"><div><h3>版本维护</h3><p>管理桌面安装包，向客户端提供匹配的更新版本。</p></div><button className="primary" disabled={busy||composing} onClick={()=>setComposing(true)}><Plus size={16}/>发布版本</button></header>
  {notice&&<p className="release-feedback" role="status"><Check size={16}/>{notice}</p>}
  <div className="release-mac-targets" aria-label="macOS 架构发布情况">{(['arm64','x64'] as const).map(arch=>{const current=latestFor(items,'darwin',arch);return <div key={arch}><span>{arch==='arm64'?'Apple 芯片':'Intel 芯片'}<small>macOS · {arch}</small></span><strong>{loading?'读取中…':error?'读取失败':current?.version??'尚未发布'}</strong></div>;})}</div>
  {composing&&<ReleaseComposer busy={busy} onUpload={upload} onCancel={()=>setComposing(false)}/>}
  <section className="release-history" aria-label="发布记录">
   <div className="release-section-head"><h4>发布记录 <span>{items.length}</span></h4><button className="release-icon-button" disabled={busy||loading} onClick={()=>void load()} aria-label="刷新版本列表"><RefreshCw size={16}/></button></div>
   <div className="release-toolbar"><nav aria-label="按操作系统筛选"><button aria-pressed={filter==='all'} onClick={()=>setFilter('all')}>全部</button>{Object.entries(releasePlatforms).map(([value,label])=><button key={value} aria-pressed={filter===value} onClick={()=>setFilter(value as ReleasePlatform)}>{label}</button>)}</nav><label><span className="sr-only">按架构筛选</span><select value={archFilter} onChange={event=>setArchFilter(event.target.value)}><option value="all">所有架构</option><option value="arm64">arm64</option><option value="x64">x64</option><option value="">未标注架构</option></select></label></div>
   {error?<div className="release-empty" role="alert"><strong>版本列表读取失败</strong><p>{error}</p><button className="release-secondary" disabled={loading} onClick={()=>void load()}>重新加载</button></div>:loading?<div className="release-loading" role="status" aria-label="正在读取版本列表"><div/><div/><div/><span className="sr-only">正在读取版本列表…</span></div>:filtered.length?<><div className="release-columns" aria-hidden="true"><span>版本</span><span>系统 / 架构</span><span>安装包</span><span>操作</span></div><ul className="release-records" aria-label="已发布版本列表">{filtered.map(item=><ReleaseRecord key={item.id} item={item} latest={latestIds.has(item.id)} busy={busy} onDelete={()=>setPending(item)}/>)}</ul></>:<div className="release-empty"><Package size={24}/><strong>{items.length?'没有匹配的安装包':'还没有发布版本'}</strong><p>{items.length?'选择其他系统或架构查看发布记录。':'分别上传 macOS 的 Apple 芯片与 Intel 安装包，客户端会按自身架构检查。'}</p><button className="release-secondary" onClick={()=>items.length?(setFilter('all'),setArchFilter('all')):setComposing(true)}>{items.length?'清除筛选':'发布第一个版本'}</button></div>}
  </section>
  {pending&&<Modal title="删除版本" onClose={()=>{if(!busy)setPending(undefined);}}><p className="organization-help">删除 {pending.version} 的 {releaseTarget(pending.platform,pending.arch)} 安装包？客户端将无法再下载此文件。</p><p>{pending.filename}</p><div className="dialog-actions"><button disabled={busy} onClick={()=>setPending(undefined)}>取消</button><button className="danger" disabled={busy} onClick={()=>run(async()=>{await cloudJson('/admin/release','DELETE',{id:pending.id});if(!alive.current)return;setPending(undefined);setNotice('版本已删除');await load();})}>确认删除</button></div></Modal>}
 </section>;
}
