import type {CommentAnchor} from './comment-model';
import {validCommentAnchor} from './comment-model';
export interface ReadingPosition {version:1;anchor:CommentAnchor|null;offset:number;scroll:number;ratio:number;updatedAt:number;}
export const READING_KEY='mydoc:reading-positions:v1';
export function validReadingPosition(value:unknown):value is ReadingPosition{const v=value as ReadingPosition;return !!v&&v.version===1&&(v.anchor===null||validCommentAnchor(v.anchor))&&Number.isFinite(v.offset)&&Math.abs(v.offset)<100000&&Number.isFinite(v.scroll)&&v.scroll>=0&&v.scroll<=10000000&&Number.isFinite(v.ratio)&&v.ratio>=0&&v.ratio<=1&&Number.isFinite(v.updatedAt)&&v.updatedAt>0;}
export const readingKey=(workspace:string,document:string,pane:string)=>JSON.stringify([workspace,document,pane]);
export function boundedScroll(value:number,max:number){return Math.max(0,Math.min(Number.isFinite(value)?value:0,Math.max(0,max)));}
export class ReadingPositions {
 constructor(private storage:{getItem:(key:string)=>string|null;setItem:(key:string,value:string)=>void|Promise<void>}){}
 private read():Record<string,ReadingPosition>{try{const value:unknown=JSON.parse(this.storage.getItem(READING_KEY)||'{}');if(!value||typeof value!=='object'||Array.isArray(value))return {};return Object.fromEntries(Object.entries(value).filter(([key,entry])=>key.length<=1000&&validReadingPosition(entry)).sort((a,b)=>(b[1] as ReadingPosition).updatedAt-(a[1] as ReadingPosition).updatedAt).slice(0,1000));}catch{return {};}}
 get(key:string){return this.read()[key];}
 put(key:string,position:ReadingPosition){if(!validReadingPosition(position)||key.length>1000)throw new Error('阅读位置无效');const entries=Object.entries({...this.read(),[key]:position}).sort((a,b)=>b[1].updatedAt-a[1].updatedAt).slice(0,1000);return this.storage.setItem(READING_KEY,JSON.stringify(Object.fromEntries(entries)));}
}

export function readingSaveError(error:unknown) {
 const name=error&&typeof error==='object'&&'name' in error?String(error.name):'';
 const message=error instanceof Error?error.message:'';
 const reason=name==='QuotaExceededError'||name==='NS_ERROR_DOM_QUOTA_REACHED'?'项目存储空间不足':name==='SecurityError'?'禁止访问项目存储':message==='阅读位置无效'?'捕获的阅读位置数据超出有效范围':`项目 SQLite 存储写入异常${message?'（'+message+'）':''}`;
 return `阅读位置保存失败：${reason}。正文内容不受影响；可在文档选项中重试保存阅读位置。`;
}
