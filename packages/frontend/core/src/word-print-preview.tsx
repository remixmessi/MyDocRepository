import {useEffect,useRef,useState} from 'react';
import {Modal} from './dialogs';
import type {WordPageLayout} from '../../../common/nbstore/src/office-word-page';
import {wordHeaderParts,wordHeaderPair,wordHeaderImageKey,WORD_HEADER_SLOTS,type WordHeaderFooter,type WordHeaderLine,type WordHeaderPart,type WordHeaderScope} from '../../../common/nbstore/src/office-word-header';
import type {WordHeaderPicture} from './office-word-page';
import type {LocalWorkspace} from './workspace';
/** CSS string literal: quotes, backslashes and line breaks escaped (margin boxes keep "\A" breaks with white-space:pre). */
const cssString=(text:string)=>'"'+text.replaceAll('\\','\\\\').replaceAll('"','\\"').replaceAll('\n','\\A ')+'"';
/**
 * CSS page margin boxes repeat headers and footers on every printed page; page fields become counters and pictures
 * (passed as `url(...)` values) sit above the slot's text. First-page and even-page (`:left`) variants override the
 * default, with :first last so it wins on page 1.
 */
export function wordMarginBoxes(value:WordHeaderFooter|undefined,pictures:Record<string,string>={}){
 if(!value)return '';const text=(line:string)=>wordHeaderParts(line).map(part=>'field' in part?(part.field==='PAGE'?'counter(page)':'counter(pages)'):cssString(part.text)).join(' ');
 const content=(scope:WordHeaderScope,part:WordHeaderPart,slot:keyof WordHeaderLine,line:string)=>{const picture=pictures[wordHeaderImageKey(scope,part,slot)];return [picture,picture&&line?'"\\A "':'',line?text(line):''].filter(Boolean).join(' ')||'""';};
 const boxes=(scope:WordHeaderScope)=>{const pair=wordHeaderPair(value,scope);if(!pair)return '';return (['header','footer'] as const).map(part=>WORD_HEADER_SLOTS.map(slot=>`@${part==='header'?'top':'bottom'}-${slot}{content:${content(scope,part,slot,pair[part][slot])};white-space:pre;font:9pt Arial,"Microsoft YaHei",sans-serif;color:#444}`).join('')).join('');};
 return `@page{${boxes('main')}}`+(value.even?`@page:left{${boxes('even')}}`:'')+(value.first?`@page:first{${boxes('first')}}`:'');
}
/** A picture as a CSS image of its stored size: an SVG wrapper fixes the size, since margin boxes cannot scale images. */
async function pictureUrls(workspace:LocalWorkspace,id:string,pictures:Record<string,WordHeaderPicture>){
 const result:Record<string,string>={};
 for(const [key,picture]of Object.entries(pictures)){
  const asset=await workspace.documentAssets(id).get(picture.assetId);if(!asset)throw new Error('页眉页脚图片缺失，请重新插入后打印');
  let binary='';const bytes=new Uint8Array(asset.bytes);for(let i=0;i<bytes.length;i+=0x8000)binary+=String.fromCharCode(...bytes.subarray(i,i+0x8000));
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${picture.width}" height="${picture.height}"><image width="${picture.width}" height="${picture.height}" preserveAspectRatio="none" href="data:${asset.mime};base64,${btoa(binary)}" xlink:href="data:${asset.mime};base64,${btoa(binary)}"/></svg>`;
  result[key]=`url("data:image/svg+xml;base64,${btoa(svg)}")`;
 }
 return result;
}
function printStyles(page:WordPageLayout){return `@page{size:${page.width/20}pt ${page.height/20}pt;margin:${page.top/20}pt ${page.right/20}pt ${page.bottom/20}pt ${page.left/20}pt}body{margin:0;color:#111;background:white;font-family:Arial,"Microsoft YaHei",sans-serif;font-size:12pt;line-height:1.6}.word-paper{width:auto!important;min-height:0!important;padding:0!important;zoom:1!important;box-shadow:none!important;margin:0!important}.word-rich-input{white-space:pre-wrap;overflow-wrap:anywhere;outline:0;border:0;min-height:1.6em}.word-cell-paragraph-break{display:block;height:0;overflow:hidden}.word-paragraph-content[data-heading="1"]{font-size:24pt;font-weight:bold}.word-paragraph-content[data-heading="2"]{font-size:18pt;font-weight:bold}.word-paragraph-content[data-heading="3"]{font-size:14pt;font-weight:bold}.word-paragraph-content[data-page-break=true]{break-before:page}.office-word-paragraph{margin:0 0 6pt;orphans:2;widows:2}table{width:100%;border-collapse:collapse}td{border:1px solid #888;vertical-align:top;padding:4pt}tr{break-inside:avoid}img{max-width:100%;object-fit:contain}.office-word-image{break-inside:avoid}a{color:inherit;text-decoration:underline}`;}
export function WordPrintPreview({source,page,headerFooter,headerImages={},workspace,id,name,onClose}:{source:HTMLElement;page:WordPageLayout;headerFooter?:WordHeaderFooter;headerImages?:Record<string,WordHeaderPicture>;workspace:LocalWorkspace;id:string;name:string;onClose:()=>void}){
 const root=useRef<HTMLDivElement>(null),snapshot=useRef<HTMLElement>(),[error,setError]=useState(''),[busy,setBusy]=useState(false);
 if(!snapshot.current){const clone=source.cloneNode(true) as HTMLElement;for(const node of clone.querySelectorAll('button,input,select,textarea,fieldset,header,.word-table-actions,caption,[role=alert],.word-page-break-marker,.word-collab-layer,.word-margin-layer,.word-image-handle,.word-image-size-tip'))node.remove();for(const node of clone.querySelectorAll('[contenteditable]'))node.removeAttribute('contenteditable');for(const node of clone.querySelectorAll('[tabindex]'))node.removeAttribute('tabindex');clone.style.zoom='1';snapshot.current=clone;}
 useEffect(()=>{root.current?.replaceChildren(snapshot.current!.cloneNode(true));},[]);
 const print=async()=>{if(busy)return;setBusy(true);setError('');const frame=document.createElement('iframe');frame.title='打印文档';frame.style.cssText='position:fixed;width:1px;height:1px;left:-10000px;top:0;border:0';document.body.append(frame);
  try{const doc=frame.contentDocument;if(!doc||!frame.contentWindow)throw new Error('无法打开打印预览');doc.title=name;const style=doc.createElement('style');style.textContent=printStyles(page)+wordMarginBoxes(headerFooter,await pictureUrls(workspace,id,headerImages));doc.head.append(style);doc.body.append(snapshot.current!.cloneNode(true));await Promise.all([...doc.images].map(image=>image.decode().catch(()=>{throw new Error('图片尚未载入，请稍后重试打印');})));await doc.fonts.ready;const cleanup=()=>frame.remove();frame.contentWindow.addEventListener('afterprint',cleanup,{once:true});setTimeout(cleanup,60000);frame.contentWindow.focus();frame.contentWindow.print();}catch(error){frame.remove();setError(error instanceof Error?error.message:'打印失败');}finally{setBusy(false);}
 };
 return <Modal title="Word 打印预览" wide onClose={onClose}><p>这是打开预览时的内容快照，后续协作修改不会改变本次打印。按当前纸张、页边距与段前分页交给浏览器分页，可在系统打印界面另存为 PDF；复杂表格及精确 Word 分页仍受浏览器排版能力限制。{headerFooter&&'页眉、页脚和页码在打印时按每页生成，预览中不显示。'}</p><div className="word-print-preview" ref={root}/>{error&&<p role="alert">{error}</p>}<div className="dialog-actions"><button disabled={busy} onClick={onClose}>关闭预览</button><button className="primary" disabled={busy} onClick={()=>void print()}>{busy?'正在准备打印…':'打印 / 另存 PDF'}</button></div></Modal>;
}
