import {validOfficeCommentAnchor,type OfficeCommentAnchor} from './office-comment-anchor';
import * as Y from 'yjs';
export const OFFICE_COMMENT_LIMITS={threads:1000,messages:10000,perThread:100,body:2000,bytes:2*1024*1024};
/** `mentions` lists @-mentioned member ids (server-checked members who can read the file). */
export interface OfficeCommentMessage{id:string;author:string;authorName:string;body:string;created:number;mentions?:string[];}
export const MAX_COMMENT_MENTIONS=20;
export const validCommentMentions=(value:unknown):value is string[]=>Array.isArray(value)&&value.length>=1&&value.length<=MAX_COMMENT_MENTIONS&&new Set(value).size===value.length&&value.every(item=>typeof item==='string'&&/^[a-zA-Z0-9-]{1,100}$/.test(item));
export interface OfficeCommentThread{anchor?:OfficeCommentAnchor;id:string;messages:OfficeCommentMessage[];resolved:boolean;updated:number;}
export type OfficeCommentOperation={operationId:string;threadId:string;base:string;type:'create'|'reply'|'resolve'|'reopen'|'delete';body?:string;anchor?:OfficeCommentAnchor;mentions?:string[]};
const record=(v:unknown):v is Record<string,unknown>=>!!v&&typeof v==='object'&&!Array.isArray(v)&&!(v instanceof Y.AbstractType);
const id=(v:unknown):v is string=>typeof v==='string'&&/^[a-zA-Z0-9-]{1,100}$/.test(v);
const fields=(v:Record<string,unknown>,names:string[])=>Object.keys(v).length===names.length&&names.every(k=>Object.hasOwn(v,k));
export function validOfficeComment(v:unknown):v is OfficeCommentThread{
 if(!record(v)||!fields(v,v.anchor===undefined?['id','messages','resolved','updated']:['id','messages','resolved','updated','anchor'])||v.anchor!==undefined&&!validOfficeCommentAnchor(v.anchor)||!id(v.id)||typeof v.resolved!=='boolean'||!Number.isSafeInteger(v.updated)||Number(v.updated)<1||!Array.isArray(v.messages)||!v.messages.length||v.messages.length>OFFICE_COMMENT_LIMITS.perThread)return false;
 const seen=new Set<string>();return v.messages.every(m=>{if(!record(m)||!fields(m,m.mentions===undefined?['id','author','authorName','body','created']:['id','author','authorName','body','created','mentions'])||m.mentions!==undefined&&!validCommentMentions(m.mentions)||!id(m.id)||seen.has(m.id)||!id(m.author)||typeof m.authorName!=='string'||!m.authorName.trim()||m.authorName.length>160||typeof m.body!=='string'||!m.body.trim()||m.body.length>OFFICE_COMMENT_LIMITS.body||!Number.isSafeInteger(m.created)||Number(m.created)<1||Number(m.created)>Number(v.updated))return false;seen.add(m.id);return true;});
}
export function officeCommentRevision(thread:OfficeCommentThread){return JSON.stringify([thread.updated,thread.messages.at(-1)?.id,thread.messages.length,thread.resolved]);}
export function validateOfficeComments(doc:Y.Doc){
 let count=0,size=0;
 if(doc.share.has('officeComments')){const comments=doc.getMap('officeComments');if(comments._start||comments.size>OFFICE_COMMENT_LIMITS.threads)throw new Error('办公评论数量或结构无效');for(const [key,value]of comments){if(!validOfficeComment(value)||key!==value.id)throw new Error('办公评论内容无效');count+=value.messages.length;size+=JSON.stringify(value).length;}if(count>OFFICE_COMMENT_LIMITS.messages||size>OFFICE_COMMENT_LIMITS.bytes)throw new Error('办公评论超过容量上限');}
 if(doc.share.has('officeCommentOperations')){const operations=doc.getMap('officeCommentOperations');if(operations._start||operations.size>10000||[...operations].some(([key,value])=>key.length>201||typeof value!=='string'||! /^[a-f0-9]{64}$/.test(value)))throw new Error('办公评论回执无效');}
}
/** Clients may carry unchanged comment state in Yjs deltas, but cannot impersonate an author. */
export function assertOfficeCommentsUnchanged(current:Uint8Array|null,update:Uint8Array){
 const before=new Y.Doc(),after=new Y.Doc();try{if(current){Y.applyUpdate(before,current);Y.applyUpdate(after,current);}Y.applyUpdate(after,update);
  for(const root of ['officeComments','officeCommentOperations']){const entries=(doc:Y.Doc)=>doc.share.has(root)?[...doc.getMap(root)].sort(([a],[b])=>a.localeCompare(b)):[];if(JSON.stringify(entries(before))!==JSON.stringify(entries(after)))throw new Error('办公评论必须通过评论接口提交');}
 }finally{before.destroy();after.destroy();}
}
