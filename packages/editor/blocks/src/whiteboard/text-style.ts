import type {Shape} from './model';
export interface TextStyle {size?:number;color?:string;bold?:boolean;italic?:boolean;align?:'left'|'center'|'right';rich?:boolean;autoHeight?:boolean}
export function validTextStyle(s:Shape){const t=s.textStyle;return t===undefined||!!t&&typeof t==='object'&&(t.size===undefined||Number.isFinite(t.size)&&t.size>=8&&t.size<=96)&&(t.color===undefined||/^#[0-9a-fA-F]{6}$/.test(t.color))&&['bold','italic','rich','autoHeight'].every(key=>t[key as 'bold']===undefined||typeof t[key as 'bold']==='boolean')&&(t.align===undefined||['left','center','right'].includes(t.align));}
export const textFont=(s:Shape)=>`${s.textStyle?.italic?'italic ':''}${s.textStyle?.bold?'bold ':''}${s.textStyle?.size??16}px system-ui`;
export const lineHeight=(s:Shape)=>(s.textStyle?.size??16)*1.5;
