import { useEffect, useRef, useState } from 'react';
import { X } from 'lucide-react';
import type { Node as PMNode } from 'prosemirror-model';
import type { WorkspaceAssets } from '@mydoc/nbstore';
import { renderDocumentExport, exportDocumentHtml, defaultExportOptions, exportFilename, type ExportOptions, type ExportSnapshot } from './document-export';
export interface ExportRequest { doc:PMNode; title:string }
export function DocumentExportDialog({request,assets,resolveTitle,onClose}:{request:ExportRequest;assets?:WorkspaceAssets;resolveTitle:(id:string,workspace:string)=>string|undefined;onClose:()=>void}){
 const dialog=useRef<HTMLDialogElement>(null),frame=useRef<HTMLIFrameElement>(null),downloadUrl=useRef('');
 const [snapshot,setSnapshot]=useState<ExportSnapshot>(),[error,setError]=useState(''),[retry,setRetry]=useState(0),[loading,setLoading]=useState(true),[ready,setReady]=useState(false),[previewWarning,setPreviewWarning]=useState('');
 const [options,setOptions]=useState<ExportOptions>(defaultExportOptions);
 useEffect(()=>{const previous=document.activeElement as HTMLElement|null;dialog.current!.showModal();return()=>{dialog.current?.close();previous?.focus();if(downloadUrl.current)URL.revokeObjectURL(downloadUrl.current);};},[]);
 useEffect(()=>{let active=true;setLoading(true);setSnapshot(undefined);setError('');setReady(false);void renderDocumentExport(request.doc,request.title,assets,resolveTitle,location.href).then(result=>{if(active)setSnapshot(result);}).catch(e=>{if(active)setError(e instanceof Error?e.message:'导出准备失败，请重试');}).finally(()=>{if(active)setLoading(false);});return()=>{active=false;};},[request,assets,retry]);
 let html='';try{if(snapshot)html=exportDocumentHtml(snapshot,options);}catch(e){html='';}
 useEffect(()=>{setReady(false);setPreviewWarning('');},[html]);
 const loaded=async()=>{const host=frame.current,doc=host?.contentDocument;if(!host||!doc||!html)return;let timeout:ReturnType<typeof setTimeout>|undefined;
  try{await Promise.race([Promise.all([doc.fonts.ready,...[...doc.images].map(image=>image.decode().catch(()=>{const label=doc.createElement('p');label.textContent='图片未加载：'+(image.alt||'图片');label.className='export-note';image.replaceWith(label);if(host.contentDocument===doc)setPreviewWarning('部分图片未加载，打印预览保留文字说明；可重新准备导出后重试。');}))]),new Promise<void>(resolve=>{timeout=setTimeout(()=>{if(host.contentDocument===doc)setPreviewWarning('部分资源加载超时，请检查预览，必要时重新准备导出。');resolve();},10000);})]);if(host.contentDocument===doc)setReady(true);}catch{if(host.contentDocument===doc)setError('打印预览加载失败，请重试');}finally{clearTimeout(timeout);}
 };
 const download=()=>{if(!html)return;try{if(downloadUrl.current)URL.revokeObjectURL(downloadUrl.current);downloadUrl.current=URL.createObjectURL(new Blob([html],{type:'text/html;charset=utf-8'}));const link=document.createElement('a');link.href=downloadUrl.current;link.download=exportFilename(request.title,'html');link.click();}catch{setError('HTML 下载失败，请重试');}};
 const print=()=>{try{if(!ready||!frame.current?.contentWindow)throw new Error();frame.current.contentWindow.focus();frame.current.contentWindow.print();}catch{setError('无法打开打印对话框，请下载 HTML 后在浏览器中打印。');}};
 return <dialog ref={dialog} className="app-dialog export-dialog" aria-label="导出文档" onCancel={event=>{event.preventDefault();onClose();}} onKeyDown={event=>{if(event.key==='Escape'&&!event.nativeEvent.isComposing){event.preventDefault();event.stopPropagation();onClose();}}}>
  <div className="dialog-heading"><h2>导出文档</h2><button type="button" aria-label="关闭导出预览" onClick={onClose}><X size={18}/></button></div>
  <p className="export-dialog-title">{request.title||'无标题文档'}</p><p className="organization-help">导出打开时的正文快照。本地文件内嵌到 HTML；打印对话框中选择“另存为 PDF”，可设置页码与页眉页脚，最终分页以打印对话框为准。</p>
  <div className="export-settings"><label>纸张<select aria-label="导出纸张" value={options.paper} onChange={e=>setOptions(o=>({...o,paper:e.target.value as ExportOptions['paper']}))}><option value="A4">A4</option><option value="Letter">Letter</option></select></label><label>方向<select aria-label="导出方向" value={options.orientation} onChange={e=>setOptions(o=>({...o,orientation:e.target.value as ExportOptions['orientation']}))}><option value="portrait">纵向</option><option value="landscape">横向</option></select></label><label>页边距<select aria-label="导出页边距" value={options.margin} onChange={e=>setOptions(o=>({...o,margin:Number(e.target.value) as ExportOptions['margin']}))}><option value={12}>紧凑 · 12mm</option><option value={20}>标准 · 20mm</option><option value={25}>宽 · 25mm</option></select></label><label className="export-title-toggle"><input type="checkbox" checked={options.includeTitle} onChange={e=>setOptions(o=>({...o,includeTitle:e.target.checked}))}/>包含文档标题</label></div>
  {loading&&<p role="status">正在准备图片、字体与图表…</p>}{error&&<div role="alert" className="form-error">{error}<button disabled={loading} onClick={()=>setRetry(n=>n+1)}>重新准备导出</button></div>}{!loading&&!error&&!html&&<p role="alert">导出文件超过大小限制，请减少文档中的文件后重试。</p>}
  {!!snapshot?.warnings.length&&<ul className="export-warnings">{snapshot.warnings.map(w=><li key={w}>{w}</li>)}</ul>}{previewWarning&&<p className="organization-warning" role="status">{previewWarning}</p>}
  {html&&<iframe ref={frame} title="文档导出预览" sandbox="allow-same-origin allow-modals allow-downloads" srcDoc={html} onLoad={()=>void loaded()}/>}
  <div className="dialog-actions"><span className="organization-help" role="status">{snapshot&&!ready&&!error?'正在加载打印预览…':''}</span><button type="button" onClick={onClose}>取消</button><button disabled={!html||loading} onClick={download}>下载 HTML</button><button className="primary" disabled={!ready||!html||loading} onClick={print}>打印 / 保存为 PDF</button></div>
 </dialog>;
}
