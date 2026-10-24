import {useEffect,useRef,useState} from 'react';
import type {PDFDocumentProxy} from 'pdfjs-dist';
import {pdfCanvasSize} from '../../../editor/blocks/src/pdf-preview-utils';
import type {PdfPageEntry} from '../../../common/nbstore/src/office-pdf';
const LIMIT=500;
function Thumbnail({pdf,entry,position,current,marks,discussions,selected,onSelect,onPick}:{pdf:PDFDocumentProxy;entry:PdfPageEntry;position:number;current:boolean;marks:number;discussions:number;selected?:boolean;onSelect?:(on:boolean,range:boolean)=>void;onPick:()=>void}){
 const ref=useRef<HTMLButtonElement>(null),canvas=useRef<HTMLCanvasElement>(null),[visible,setVisible]=useState(false),[ready,setReady]=useState(false),source='source' in entry?entry.source:undefined,rotate=entry.rotate??0;
 useEffect(()=>{const observer=new IntersectionObserver(entries=>{if(entries[0].isIntersecting)setVisible(true);},{rootMargin:'200px'});if(ref.current)observer.observe(ref.current);return()=>observer.disconnect();},[]);
 useEffect(()=>{if(current)ref.current?.scrollIntoView({block:'nearest'});},[current]);
 useEffect(()=>{if(!visible||!source||source>pdf.numPages)return;let active=true,task:ReturnType<import('pdfjs-dist').PDFPageProxy['render']>|undefined;setReady(false);void(async()=>{const page=await pdf.getPage(source);try{const angle=(page.rotate+rotate)%360,natural=page.getViewport({scale:1,rotation:angle}),scale=Math.min(120/natural.width,1),view=page.getViewport({scale,rotation:angle}),size=pdfCanvasSize(natural.width,natural.height,scale,window.devicePixelRatio||1),element=canvas.current;if(!element||!active)return;element.width=size.width;element.height=size.height;element.style.width=view.width+'px';element.style.height=view.height+'px';task=page.render({canvas:element,viewport:view,transform:[size.ratio,0,0,size.ratio,0,0]});await task.promise;if(active)setReady(true);}catch{}finally{page.cleanup();}})();return()=>{active=false;task?.cancel();};},[visible,pdf,source,rotate]);
 const blank='blank' in entry?entry.blank:undefined,turned=rotate%180!==0;
 return <div className="pdf-thumbnail-item" data-selected={selected||undefined}>{onSelect&&<input type="checkbox" className="pdf-thumbnail-check" aria-label={`选择第 ${position} 页`} checked={!!selected} onChange={event=>onSelect(event.target.checked,(event.nativeEvent as MouseEvent).shiftKey)}/>}
  <button ref={ref} className="pdf-thumbnail" aria-label={`第 ${position} 页`} aria-current={current?'page':undefined} onClick={event=>{if(onSelect&&(event.ctrlKey||event.metaKey||event.shiftKey)){onSelect(!selected,event.shiftKey);return;}onPick();}}>
   <span className="pdf-thumbnail-page">{blank?<span className="pdf-thumbnail-blank" style={{aspectRatio:`${turned?blank.height:blank.width} / ${turned?blank.width:blank.height}`}}/>:<canvas ref={canvas} aria-hidden="true" style={{visibility:ready?'visible':'hidden'}}/>}</span>
   <span>{position}{blank?' · 空白页':''}{rotate?` · ${rotate}°`:''}{marks?` · ${marks} 条批注`:''}</span>{discussions>0&&<span className="pdf-thumbnail-discussions" title={`${discussions} 条未解决讨论`}>{discussions} 条讨论</span>}</button></div>;
}
/**
 * Left-hand page strip like the WPS/Acrobat page panel, in the arranged order; pages render only when scrolled into
 * view. With `onSelect`, checkboxes (and Ctrl/Shift+click) select pages for 页面管理.
 */
export function PdfThumbnails({pdf,pages,position,counts,discussions,selected,onSelect,onPosition}:{pdf:PDFDocumentProxy;pages:PdfPageEntry[];position:number;counts:Map<number,number>;discussions?:Map<number,number>;selected?:Set<string>;onSelect?:(id:string,on:boolean,range:boolean)=>void;onPosition:(position:number)=>void}){
 const shown=pages.slice(0,LIMIT);
 return <div className="pdf-thumbnails" role="navigation" aria-label="页面缩略图">{shown.map((entry,i)=>{const source='source' in entry?entry.source:0;return <Thumbnail key={entry.id} pdf={pdf} entry={entry} position={i+1} current={position===i+1} marks={source?counts.get(source)??0:0} discussions={source?discussions?.get(source)??0:0} selected={selected?.has(entry.id)} onSelect={onSelect?(on,range)=>onSelect(entry.id,on,range):undefined} onPick={()=>onPosition(i+1)}/>;})}{pages.length>LIMIT&&<p>仅显示前 {LIMIT} 页缩略图，可在状态栏输入页码跳转。</p>}</div>;
}
