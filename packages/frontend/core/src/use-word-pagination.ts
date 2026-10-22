import {useEffect,useState,type RefObject} from 'react';
import type {WordPageLayout} from '../../../common/nbstore/src/office-word-page';
export type WordPageBreak={top:number;page:number};
/**
 * Automatic pages from the browser's own line layout: page body height comes from paper size and margins,
 * manual "page break before" starts a page, images and table rows move whole, and text breaks between lines.
 * This mirrors how printing will paginate; it is not Word's layout engine.
 */
export function paginateWordBlocks(blocks:{top:number;height:number;lines:number[];breakBefore:boolean;whole:boolean}[],body:number){
 const breaks:WordPageBreak[]=[];let pageTop=0,page=1;
 const start=(top:number)=>{pageTop=top;page++;breaks.push({top,page});};
 for(const block of blocks){
  if(block.breakBefore&&block.top>pageTop)start(block.top);
  if(block.top+block.height<=pageTop+body)continue;
  if(block.whole||!block.lines.length){if(block.top>pageTop)start(block.top);continue;}
  // Line boxes (offsets of each line's top within the block) let text continue on the next page.
  for(const [i,line]of block.lines.entries()){const top=block.top+line,bottom=block.top+(block.lines[i+1]??block.height);if(bottom>pageTop+body&&top>pageTop)start(top);}
 }
 return breaks;
}
const PX_PER_TWIP=96/1440;
export function useWordPagination(paper:RefObject<HTMLElement>,page:WordPageLayout,enabled=true){
 const [breaks,setBreaks]=useState<WordPageBreak[]>([]);
 useEffect(()=>{
  const element=paper.current;if(!element||!enabled)return;let frame=0;
  const measure=()=>{cancelAnimationFrame(frame);frame=requestAnimationFrame(()=>{
   const body=(page.height-page.top-page.bottom)*PX_PER_TWIP,origin=element.getBoundingClientRect(),zoom=origin.width/(element.offsetWidth||1)||1,top=page.top*PX_PER_TWIP;
   const blocks=[...element.children].filter((child):child is HTMLElement=>child instanceof HTMLElement&&!child.classList.contains('word-page-break-marker')).flatMap(child=>{
    const rect=child.getBoundingClientRect();if(!rect.height)return [];const blockTop=(rect.top-origin.top)/zoom-top,input=child.querySelector<HTMLElement>(':scope .word-paragraph-content .word-rich-input'),whole=!input||child.classList.contains('office-word-image');
    let lines:number[]=[];if(input&&!whole){const range=document.createRange();range.selectNodeContents(input);const seen=new Set<number>();for(const box of range.getClientRects()){const offset=Math.round((box.top-rect.top)/zoom);if(offset>=0&&!seen.has(offset)){seen.add(offset);lines.push(offset);}}lines=lines.sort((a,b)=>a-b);if(lines[0]!==0)lines.unshift(0);}
    if(child.classList.contains('office-word-table'))return [...child.querySelectorAll<HTMLElement>('tr')].map(row=>{const box=row.getBoundingClientRect();return {top:(box.top-origin.top)/zoom-top,height:box.height/zoom,lines:[],breakBefore:false,whole:true};});
    return [{top:blockTop,height:rect.height/zoom,lines,breakBefore:!!child.querySelector('[data-page-break="true"]'),whole}];
   });
   const next=paginateWordBlocks(blocks,body).map(entry=>({...entry,top:entry.top+top}));setBreaks(previous=>JSON.stringify(previous)===JSON.stringify(next)?previous:next);
  });};
  const resize=new ResizeObserver(measure),mutation=new MutationObserver(measure);resize.observe(element);mutation.observe(element,{childList:true,subtree:true,characterData:true,attributes:true,attributeFilter:['data-page-break','style']});measure();
  return()=>{cancelAnimationFrame(frame);resize.disconnect();mutation.disconnect();};
 },[paper,page.width,page.height,page.top,page.bottom,page.left,page.right,enabled]);
 return breaks;
}
