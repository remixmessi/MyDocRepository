import { useEffect, useRef, useState } from 'react';
import type { PDFDocumentLoadingTask, PDFDocumentProxy } from 'pdfjs-dist';
import type { AssetRecord, WorkspaceAssets } from '@mydoc/nbstore';
import { pdfCanvasSize, pdfPageNumber } from './pdf-preview-utils';

export function PdfPreview({assetId,assets,onClose}:{assetId:string;assets:WorkspaceAssets;onClose:()=>void}) {
 const dialog=useRef<HTMLDialogElement>(null),stage=useRef<HTMLDivElement>(null),passwordCallback=useRef<(password:string)=>void>();
 const [asset,setAsset]=useState<AssetRecord>(),[pdf,setPdf]=useState<PDFDocumentProxy>(),[retry,setRetry]=useState(0),[error,setError]=useState(''),[loading,setLoading]=useState(true);
 const [page,setPage]=useState(1),[pageInput,setPageInput]=useState('1'),[zoom,setZoom]=useState('fit'),[width,setWidth]=useState(600),[rendering,setRendering]=useState(false),[pageError,setPageError]=useState(''),[pageRetry,setPageRetry]=useState(0);
 const [passwordPrompt,setPasswordPrompt]=useState(''),[password,setPassword]=useState('');
 useEffect(()=>{const previous=document.activeElement as HTMLElement|null;dialog.current!.showModal();return()=>{dialog.current?.close();previous?.isConnected&&previous.focus();};},[]);
 useEffect(()=>{const element=stage.current!;const observer=new ResizeObserver(entries=>setWidth(Math.max(120,entries[0].contentRect.width-32)));observer.observe(element);return()=>observer.disconnect();},[]);
 useEffect(()=>{
  let active=true,task:PDFDocumentLoadingTask|undefined;
  setLoading(true);setError('');setPdf(undefined);setAsset(undefined);setPasswordPrompt('');setPassword('');passwordCallback.current=undefined;setPage(1);setPageInput('1');
  void (async()=>{try {
   const record=await assets.get(assetId);if(!active)return;if(!record)throw new Error('本地文件缺失，请恢复包含文件的工作区备份');setAsset(record);
   const engine=await import('./pdf-engine');if(!active)return;task=engine.loadPdf(record.bytes);
   task.onPassword=(callback:(password:string)=>void,reason:number)=>{if(!active)return;passwordCallback.current=callback;setPassword('');setPasswordPrompt(reason===2?'密码不正确，请重新输入':'此 PDF 已加密，请输入密码');};
   const document=await task.promise;if(active){setPasswordPrompt('');setPdf(document);}
  }catch(e){if(active){const name=e instanceof Error?e.name:'';setError(name==='InvalidPDFException'?'PDF 文件损坏或格式无效，可下载原文件检查':e instanceof Error&&e.message.includes('本地文件缺失')?e.message:'PDF 无法读取，请重试或下载原文件');}}finally{if(active)setLoading(false);}})();
  return()=>{active=false;passwordCallback.current=undefined;void task?.destroy().catch(()=>{});};
 },[assetId,assets,retry]);
 useEffect(()=>{
  const element=stage.current!;element.querySelector('canvas')?.remove();if(!pdf)return;
  let active=true,renderTask:ReturnType<Awaited<ReturnType<PDFDocumentProxy['getPage']>>['render']>|undefined;
  setRendering(true);setPageError('');
  void (async()=>{let current:Awaited<ReturnType<PDFDocumentProxy['getPage']>>|undefined;try{
   current=await pdf.getPage(page);if(!active)return;
   const natural=current.getViewport({scale:1}),scale=zoom==='fit'?Math.min(2,width/natural.width):Number(zoom);
   const size=pdfCanvasSize(natural.width,natural.height,scale,window.devicePixelRatio||1),canvas=document.createElement('canvas');canvas.width=size.width;canvas.height=size.height;canvas.style.width=size.cssWidth+'px';canvas.style.height=size.cssHeight+'px';canvas.setAttribute('role','img');canvas.setAttribute('aria-label',`PDF 第 ${page} 页`);
   renderTask=current.render({canvas,viewport:current.getViewport({scale}),transform:[size.ratio,0,0,size.ratio,0,0]});await renderTask.promise;
   if(active){element.append(canvas);element.scrollTo(0,0);}
  }catch(e){if(active&&!(e instanceof Error&&e.name==='RenderingCancelledException'))setPageError('此页无法显示，可重试或下载原文件');}finally{current?.cleanup();if(active)setRendering(false);}})();
  return()=>{active=false;renderTask?.cancel();};
 },[pdf,page,zoom,width,pageRetry]);
 const go=(number:number)=>{setPage(number);setPageInput(String(number));};
 const jump=()=>{if(!pdf)return;const number=pdfPageNumber(pageInput,pdf.numPages);if(number===null){setPageInput(String(page));return;}go(number);};
 const fileUrl=()=>URL.createObjectURL(new Blob([asset!.bytes],{type:'application/pdf'}));
 const download=()=>{if(!asset)return;const url=fileUrl(),link=document.createElement('a');link.href=url;link.download=asset.name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
 const open=()=>{if(!asset)return;const url=fileUrl();window.open(url,'_blank','noopener,noreferrer');setTimeout(()=>URL.revokeObjectURL(url),60000);};
 return <dialog ref={dialog} className="app-dialog pdf-dialog" aria-label="PDF 附件预览" onCancel={event=>{event.preventDefault();onClose();}} onKeyDown={event=>{if(event.nativeEvent.isComposing)return;if(event.key==='Escape'){event.preventDefault();event.stopPropagation();onClose();}}}>
  <div className="dialog-heading"><h2>PDF 附件预览</h2><button type="button" aria-label="关闭 PDF 预览" onClick={onClose}>×</button></div>
  <p className="pdf-file-title">{asset?.name||'PDF 附件'}</p>
  <div className="pdf-toolbar" role="group" aria-label="PDF 阅读工具">
   <button disabled={!pdf||page<=1} onClick={()=>go(page-1)}>上一页</button><label>页码<input aria-label="PDF 页码" inputMode="numeric" value={pageInput} disabled={!pdf} onChange={event=>setPageInput(event.target.value)} onBlur={jump} onKeyDown={event=>{if(event.key==='Enter'&&!event.nativeEvent.isComposing){event.preventDefault();jump();}}}/></label><span> / {pdf?.numPages??'—'}</span><button disabled={!pdf||page>=pdf.numPages} onClick={()=>go(page+1)}>下一页</button>
   <label>缩放<select aria-label="PDF 缩放" value={zoom} disabled={!pdf} onChange={event=>setZoom(event.target.value)}><option value="fit">适合宽度</option>{[0.5,0.75,1,1.25,1.5,2].map(value=><option key={value} value={value}>{value*100}%</option>)}</select></label>
   <button disabled={!asset} onClick={download}>下载原文件</button><button disabled={!asset} onClick={open}>新窗口打开</button>
  </div>
  {passwordPrompt&&<form className="pdf-password" onSubmit={event=>{event.preventDefault();if(password){passwordCallback.current?.(password);setPasswordPrompt('正在验证密码…');}}}><label>{passwordPrompt}<input aria-label="PDF 密码" type="password" autoComplete="off" value={password} onChange={event=>setPassword(event.target.value)}/></label><button disabled={!password}>解锁</button></form>}
  {error&&<p role="alert">{error} <button onClick={()=>setRetry(value=>value+1)}>重试预览</button></p>}
  <p className="pdf-status" role="status">{loading&&!passwordPrompt?'正在读取 PDF…':rendering?'正在显示页面…':pdf?`第 ${page} 页，共 ${pdf.numPages} 页`:''}</p>
  {pageError&&<p role="alert">{pageError} <button onClick={()=>setPageRetry(value=>value+1)}>重试此页</button></p>}
  <div ref={stage} className="pdf-page-stage" tabIndex={0} aria-label="PDF 页面" aria-busy={loading||rendering} />
 </dialog>;
}
