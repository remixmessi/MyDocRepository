import {useEffect,useRef,useState} from 'react';
import type {PDFDocumentProxy} from 'pdfjs-dist';
import {pdfCanvasSize} from '../../../editor/blocks/src/pdf-preview-utils';
const LIMIT=500;
function Thumbnail({pdf,number,current,marks,discussions,onPick}:{pdf:PDFDocumentProxy;number:number;current:boolean;marks:number;discussions:number;onPick:()=>void}){
 const ref=useRef<HTMLButtonElement>(null),canvas=useRef<HTMLCanvasElement>(null),[visible,setVisible]=useState(false),[ready,setReady]=useState(false);
 useEffect(()=>{const observer=new IntersectionObserver(entries=>{if(entries[0].isIntersecting)setVisible(true);},{rootMargin:'200px'});if(ref.current)observer.observe(ref.current);return()=>observer.disconnect();},[]);
 useEffect(()=>{if(current)ref.current?.scrollIntoView({block:'nearest'});},[current]);
 useEffect(()=>{if(!visible)return;let active=true,task:ReturnType<import('pdfjs-dist').PDFPageProxy['render']>|undefined;void(async()=>{const page=await pdf.getPage(number);try{const natural=page.getViewport({scale:1}),scale=Math.min(120/natural.width,1),view=page.getViewport({scale}),size=pdfCanvasSize(natural.width,natural.height,scale,window.devicePixelRatio||1),element=canvas.current;if(!element||!active)return;element.width=size.width;element.height=size.height;element.style.width=view.width+'px';element.style.height=view.height+'px';task=page.render({canvas:element,viewport:view,transform:[size.ratio,0,0,size.ratio,0,0]});await task.promise;if(active)setReady(true);}catch{}finally{page.cleanup();}})();return()=>{active=false;task?.cancel();};},[visible,pdf,number]);
 return <button ref={ref} className="pdf-thumbnail" aria-label={`第 ${number} 页`} aria-current={current?'page':undefined} onClick={onPick}><span className="pdf-thumbnail-page"><canvas ref={canvas} aria-hidden="true" style={{visibility:ready?'visible':'hidden'}}/></span><span>{number}{marks?` · ${marks} 条批注`:''}</span>{discussions>0&&<span className="pdf-thumbnail-discussions" title={`${discussions} 条未解决讨论`}>{discussions} 条讨论</span>}</button>;
}
/** Left-hand page strip like the WPS/Acrobat page panel; pages render only when scrolled into view. */
export function PdfThumbnails({pdf,page,counts,discussions,onPage}:{pdf:PDFDocumentProxy;page:number;counts:Map<number,number>;discussions?:Map<number,number>;onPage:(page:number)=>void}){
 const total=Math.min(pdf.numPages,LIMIT);
 return <div className="pdf-thumbnails" role="navigation" aria-label="页面缩略图">{Array.from({length:total},(_,i)=><Thumbnail key={i} pdf={pdf} number={i+1} current={page===i+1} marks={counts.get(i+1)??0} discussions={discussions?.get(i+1)??0} onPick={()=>onPage(i+1)}/>)}{pdf.numPages>LIMIT&&<p>仅显示前 {LIMIT} 页缩略图，可在状态栏输入页码跳转。</p>}</div>;
}
