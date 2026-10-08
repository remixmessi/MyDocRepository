import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import type {BlockStore} from '@mydoc/editor-framework';
import type {Node as PMNode} from 'prosemirror-model';
import {AllSelection,TextSelection,type EditorState} from 'prosemirror-state';
import {absolutePositionToRelativePosition,relativePositionToAbsolutePosition,type ProsemirrorBinding} from 'y-prosemirror';
export type CommentAnchor=ReturnType<typeof Y.relativePositionToJSON>;
export interface CommentMessage {id:string;text:string;createdAt:number;updatedAt:number;}
export interface CommentThread {id:string;quote:string;from:CommentAnchor;to:CommentAnchor;resolved:boolean;createdAt:number;updatedAt:number;messages:CommentMessage[];}
export type CommentMapping=ProsemirrorBinding['mapping'];
export const MAX_COMMENT_THREADS=500,MAX_COMMENT_MESSAGES=100,MAX_COMMENT_TEXT=4000;
function record(value:unknown):value is Record<string,unknown>{return !!value&&typeof value==='object'&&!Array.isArray(value);}
function id(value:unknown):value is string{return typeof value==='string'&&/^[\w-]{1,80}$/.test(value);}
function time(value:unknown):value is number{return typeof value==='number'&&Number.isFinite(value)&&value>0;}
export function validCommentAnchor(value:unknown){if(!record(value)||typeof value.assoc!=='number'||!Number.isInteger(value.assoc)||![-1,0].includes(value.assoc))return false;const validId=(v:unknown)=>record(v)&&Number.isSafeInteger(v.client)&&Number(v.client)>=0&&Number.isSafeInteger(v.clock)&&Number(v.clock)>=0;return (value.type===undefined||validId(value.type))&&(value.item===undefined||validId(value.item))&&(value.tname===undefined||typeof value.tname==='string'&&value.tname.length<=100)&&!!(value.type||value.tname);}
export function validCommentThread(value:unknown):value is CommentThread {if(!record(value)||!id(value.id)||typeof value.quote!=='string'||!value.quote.trim()||value.quote.length>1000||!validCommentAnchor(value.from)||!validCommentAnchor(value.to)||typeof value.resolved!=='boolean'||!time(value.createdAt)||!time(value.updatedAt)||!Array.isArray(value.messages)||!value.messages.length||value.messages.length>MAX_COMMENT_MESSAGES)return false;const ids=new Set<string>();return value.messages.every(m=>{if(!record(m)||!id(m.id)||ids.has(m.id)||typeof m.text!=='string'||!m.text.trim()||m.text.length>MAX_COMMENT_TEXT||!time(m.createdAt)||!time(m.updatedAt))return false;ids.add(m.id);return true;});}
export function commentSelectionAllowed(state:EditorState){const selection=state.selection;if(!(selection instanceof TextSelection)&&!(selection instanceof AllSelection)||selection.empty||selection.to-selection.from>20000)return false;let hasText=false,code=false;state.doc.nodesBetween(selection.from,selection.to,node=>{if(node.type.spec.code)code=true;if(node.isText)hasText=true;});return hasText&&!code&&!!state.doc.textBetween(selection.from,selection.to,'\n','').trim();}
export function createCommentAnchor(doc:Y.Doc,fragment:Y.XmlFragment,mapping:CommentMapping,position:number,assoc:0|-1){const relative=absolutePositionToRelativePosition(position,fragment,mapping),absolute=Y.createAbsolutePositionFromRelativePosition(relative,doc);if(!absolute)throw new Error('选区位置已失效，请重新选择文字');return Y.relativePositionToJSON(Y.createRelativePositionFromTypeIndex(absolute.type,absolute.index,assoc));}
export function resolveComment(thread:CommentThread,doc:Y.Doc,fragment:Y.XmlFragment,mapping:CommentMapping,body:PMNode){try{const from=relativePositionToAbsolutePosition(doc,fragment,Y.createRelativePositionFromJSON(thread.from),mapping),to=relativePositionToAbsolutePosition(doc,fragment,Y.createRelativePositionFromJSON(thread.to),mapping);if(from===null||to===null||from>=to||from<0||to>body.content.size||!body.textBetween(from,to,'\n','').trim())return null;return {from,to};}catch{return null;}}
export class CommentStore {
 readonly threads:Y.Map<CommentThread>;
 constructor(readonly store:BlockStore){this.threads=store.doc.getMap<CommentThread>('documentComments');store.history.addToScope(this.threads);}
 list(){return [...this.threads.values()].filter(validCommentThread).sort((a,b)=>a.createdAt-b.createdAt||a.id.localeCompare(b.id));}
 private commit(id:string,next?:CommentThread){if(next&&!validCommentThread(next))throw new Error('评论内容或引用信息无效');this.store.history.stopCapturing();this.store.doc.transact(()=>next?this.threads.set(id,next):this.threads.delete(id),LOCAL_ORIGIN);this.store.history.stopCapturing();}
 create(quote:string,from:CommentAnchor,to:CommentAnchor,text:string){if(this.threads.size>=MAX_COMMENT_THREADS)throw new Error('每篇最多 500 条评论');const now=Date.now(),thread:CommentThread={id:crypto.randomUUID(),quote:quote.trim().slice(0,1000),from,to,resolved:false,createdAt:now,updatedAt:now,messages:[{id:crypto.randomUUID(),text:text.trim(),createdAt:now,updatedAt:now}]};this.commit(thread.id,thread);return thread.id;}
 update(id:string,change:(thread:CommentThread)=>CommentThread){const thread=this.threads.get(id);if(!thread)throw new Error('评论已删除，请重新选择');this.commit(id,{...change(thread),updatedAt:Date.now()});}
 reply(id:string,text:string){this.update(id,thread=>{if(thread.resolved)throw new Error('请先重新打开评论再回复');if(thread.messages.length>=MAX_COMMENT_MESSAGES)throw new Error('每条评论最多 100 条消息');const now=Date.now();return {...thread,messages:[...thread.messages,{id:crypto.randomUUID(),text:text.trim(),createdAt:now,updatedAt:now}]};});}
 edit(id:string,messageId:string,text:string){this.update(id,thread=>{if(!thread.messages.some(m=>m.id===messageId))throw new Error('消息已删除');return {...thread,messages:thread.messages.map(m=>m.id===messageId?{...m,text:text.trim(),updatedAt:Date.now()}:m)};});}
 deleteMessage(id:string,messageId:string){this.update(id,thread=>{if(thread.messages[0].id===messageId)throw new Error('首条评论请使用删除评论');return {...thread,messages:thread.messages.filter(m=>m.id!==messageId)};});}
 resolve(id:string,resolved:boolean){this.update(id,thread=>({...thread,resolved}));}
 remove(id:string){if(!this.threads.has(id))throw new Error('评论已删除');this.commit(id);}
}
