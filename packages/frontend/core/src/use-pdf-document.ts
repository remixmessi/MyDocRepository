import {useEffect,useState} from 'react';
import type {PDFDocumentProxy,PDFDocumentLoadingTask} from 'pdfjs-dist';
import {getCloudOfficeOriginal} from '../../../common/nbstore/src/cloud';
/** Load the immutable original once; the page view, thumbnails and search share this document. */
export function usePdfDocument(workspace:string,id:string){
 const [pdf,setPdf]=useState<PDFDocumentProxy>(),[error,setError]=useState(''),[loading,setLoading]=useState(true),[retry,setRetry]=useState(0);
 useEffect(()=>{let active=true,encrypted=false,task:PDFDocumentLoadingTask|undefined;setLoading(true);setError('');setPdf(undefined);void(async()=>{try{
  const bytes=new Uint8Array(await(await getCloudOfficeOriginal(workspace,id)).arrayBuffer()),engine=await import('../../../editor/blocks/src/pdf-engine');if(!active)return;
  task=engine.loadPdf(bytes);task.onPassword=()=>{encrypted=true;if(active){setError('加密 PDF 请下载原件查看；当前不提供解密编辑。');setLoading(false);}void task?.destroy().catch(()=>{});};
  const document=await task.promise;if(active)setPdf(document);
 }catch(error){if(active&&!encrypted)setError(error instanceof Error?'PDF 无法读取：'+error.message:'PDF 无法读取');}finally{if(active)setLoading(false);}})();return()=>{active=false;void task?.destroy().catch(()=>{});};},[workspace,id,retry]);
 return {pdf,error,loading,reload:()=>setRetry(value=>value+1)};
}
/** Find text across pages (bounded), returning page numbers with a short excerpt around each hit. */
export async function searchPdf(pdf:PDFDocumentProxy,query:string,signal:{cancelled:boolean}){
 const needle=query.trim().toLowerCase(),hits:{page:number;excerpt:string}[]=[];if(!needle)return hits;
 for(let number=1;number<=Math.min(pdf.numPages,500)&&hits.length<200;number++){
  if(signal.cancelled)return hits;const page=await pdf.getPage(number);try{const content=await page.getTextContent(),text=content.items.map(item=>'str' in item?item.str:'').join(' '),lower=text.toLowerCase();let from=0,index:number;
   while((index=lower.indexOf(needle,from))>=0&&hits.length<200){const start=Math.max(0,index-20),end=index+needle.length+30;hits.push({page:number,excerpt:(start>0?'…':'')+text.slice(start,end).trim()+(end<text.length?'…':'')});from=index+needle.length;}}finally{page.cleanup();}
 }
 return hits;
}
