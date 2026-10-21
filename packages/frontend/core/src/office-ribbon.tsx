import {useEffect,useId,useRef,useState,type ReactNode,type ButtonHTMLAttributes} from 'react';
import type {LucideIcon} from 'lucide-react';
/** Shared Office-style chrome: title bar, ribbon tabs with arrow-key navigation, labelled groups and status bar. */
export type OfficeAppKind='word'|'sheet'|'slides'|'pdf';
const BADGE:Record<OfficeAppKind,string>={word:'W',sheet:'S',slides:'P',pdf:'PDF'};
export function OfficeTitleBar({kind,name,status,children}:{kind:OfficeAppKind;name:string;status:string;children?:ReactNode}){
 return <header className="office-document-title" data-app={kind}><div className="office-file-badge" aria-hidden="true">{BADGE[kind]}</div><div><h2 title={name}>{name}</h2><span>{status}</span></div><div className="office-title-actions">{children}</div></header>;
}
export function OfficeRibbon({tabs,tab,onTab,label,children}:{tabs:string[];tab:string;onTab:(tab:string)=>void;label:string;children:ReactNode}){
 const id=useId();
 return <><div className="office-ribbon-tabs" role="tablist" aria-label={label}>{tabs.map((name,index)=><button key={name} role="tab" id={id+'-'+index} aria-selected={tab===name} aria-controls={id+'-panel'} tabIndex={tab===name?0:-1} onClick={()=>onTab(name)} onKeyDown={event=>{
   if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return;event.preventDefault();
   const next=event.key==='Home'?0:event.key==='End'?tabs.length-1:(index+(event.key==='ArrowRight'?1:tabs.length-1))%tabs.length;onTab(tabs[next]);(event.currentTarget.parentElement?.children[next] as HTMLElement|undefined)?.focus();
  }}>{name}</button>)}</div>
  <div id={id+'-panel'} className="office-ribbon-panel" role="tabpanel" aria-labelledby={id+'-'+tabs.indexOf(tab)} data-tab={tab}><div className="office-ribbon-groups">{children}</div></div></>;
}
export function RibbonGroup({label,children}:{label:string;children:ReactNode}){
 return <div className="ribbon-group" role="group" aria-label={label}><div className="ribbon-control-row">{children}</div><small aria-hidden="true">{label}</small></div>;
}
/** Icon plus text; `compact` keeps the text for assistive tech and the tooltip only, like Office's small toolbar icons. */
export function RibbonButton({icon:Icon,label,pressed,large,compact,...props}:{icon?:LucideIcon;label:string;pressed?:boolean;large?:boolean;compact?:boolean}&Omit<ButtonHTMLAttributes<HTMLButtonElement>,'children'>){
 return <button type="button" className={'ribbon-button'+(large?' ribbon-button-large':'')+(compact&&Icon?' ribbon-button-compact':'')} aria-pressed={pressed} title={compact?label:undefined} {...props}>{Icon&&<Icon size={large?20:15} aria-hidden="true"/>}<span className={compact&&Icon?'sr-only':undefined}>{label}</span></button>;
}
/** A title-bar menu (like WPS “文件/导出”); closes on outside click or Escape and returns focus to its button. */
export function OfficeMenu({label,icon:Icon,children}:{label:string;icon?:LucideIcon;children:(close:()=>void)=>ReactNode}){
 const [open,setOpen]=useState(false),root=useRef<HTMLDivElement>(null),button=useRef<HTMLButtonElement>(null);
 useEffect(()=>{if(!open)return;const outside=(event:MouseEvent)=>{if(!root.current?.contains(event.target as Node))setOpen(false);},key=(event:KeyboardEvent)=>{if(event.key==='Escape'){setOpen(false);button.current?.focus();}};document.addEventListener('mousedown',outside);document.addEventListener('keydown',key);return()=>{document.removeEventListener('mousedown',outside);document.removeEventListener('keydown',key);};},[open]);
 return <div className="office-menu" ref={root}><button ref={button} type="button" aria-haspopup="menu" aria-expanded={open} onClick={()=>setOpen(value=>!value)}>{Icon&&<Icon size={15} aria-hidden="true"/>}{label}</button>{open&&<div className="office-menu-list" role="menu" aria-label={label}>{children(()=>setOpen(false))}</div>}</div>;
}
export function OfficeStatusBar({children}:{children:ReactNode}){return <footer className="office-document-status">{children}</footer>;}
