import {validOfficeCommentAnchor,officeAnchorTarget} from '../../../../common/nbstore/src/office-comment-anchor';
import * as Y from 'yjs';
import {createHash} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {validateOfficeComments,officeCommentRevision,validCommentMentions,type OfficeCommentThread,type OfficeCommentOperation} from '../../../../common/nbstore/src/office-comments';
export type CommentNotice={user:string;kind:'mention'|'reply';thread:string;excerpt:string};
export class OfficeCommentError extends Error{constructor(message:string,readonly status=400){super(message);}}
export function submitOfficeComment(database:SQLiteProjectDatabase,workspace:string,id:string,actor:{id:string;name:string;admin:boolean},input:OfficeCommentOperation){
 if(input?.mentions!==undefined&&(!['create','reply'].includes(input.type)||!validCommentMentions(input.mentions)))throw new OfficeCommentError('提及的成员无效');
 if(!input||Object.keys(input).some(key=>!['operationId','threadId','base','type','body','anchor','mentions'].includes(key))||! /^[a-zA-Z0-9-]{16,80}$/.test(input.operationId??'')||! /^[a-zA-Z0-9-]{1,100}$/.test(input.threadId??'')||typeof input.base!=='string'||input.base.length>500||!['create','reply','resolve','reopen','delete'].includes(input.type)||(['create','reply'].includes(input.type)?typeof input.body!=='string'||!input.body.trim()||input.body.length>2000:input.body!==undefined))throw new OfficeCommentError('办公评论操作无效');
 if(input.anchor!==undefined&&(input.type!=='create'||!validOfficeCommentAnchor(input.anchor)))throw new OfficeCommentError('评论定位无效');
 return database.db.transaction(()=>{
  const scope='cloud:'+workspace,current=database.load(scope,id);if(!current)throw new OfficeCommentError('文件不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);validateOfficeComments(doc);const comments=doc.getMap<OfficeCommentThread>('officeComments'),receipts=doc.getMap<string>('officeCommentOperations'),key=actor.id+':'+input.operationId,signature=createHash('sha256').update(JSON.stringify([input.type,input.threadId,input.base,input.body??null,...(input.anchor?[input.anchor]:[]),...(input.mentions?[input.mentions]:[])])).digest('hex');
   if(receipts.has(key)){if(receipts.get(key)!==signature)throw new OfficeCommentError('评论操作编号已用于其他内容',409);return {update:current,repeated:true,notify:[] as CommentNotice[]};}if(receipts.size>=10000)throw new OfficeCommentError('评论操作记录已满，请备份恢复为新文件后继续',409);
   const thread=comments.get(input.threadId);if(input.type==='create'){if(thread||input.base!=='')throw new OfficeCommentError('评论已存在，请刷新后重试',409);}else if(!thread||officeCommentRevision(thread)!==input.base)throw new OfficeCommentError('评论已有新回复或状态变化，输入已保留，请载入最新评论后重试',409);
   if(input.type==='create'&&comments.size>=1000)throw new OfficeCommentError('每个文件最多 1000 个评论线程，请先清理已解决的讨论',413);
   if(input.type==='reply'&&thread!.messages.length>=100)throw new OfficeCommentError('单个评论线程最多 100 条消息，请发起新讨论',413);
   if(input.type==='delete'&&!actor.admin&&thread!.messages[0].author!==actor.id)throw new OfficeCommentError('只有评论发起者或管理员可删除线程',403);
   if(input.type==='reply'&&thread!.resolved)throw new OfficeCommentError('评论已解决，请先重新打开',409);
   if(input.type==='create'&&input.anchor&&!officeAnchorTarget(doc,input.anchor))throw new OfficeCommentError('评论所选内容已删除或定位失效，请重新选择',409);
   const vector=Y.encodeStateVector(doc),now=Math.max(Date.now(),(thread?.updated??0)+1);
   if(input.type==='delete')comments.delete(input.threadId);
   else if(input.type==='resolve'||input.type==='reopen')comments.set(input.threadId,{...thread!,resolved:input.type==='resolve',updated:now});
   else comments.set(input.threadId,{...(thread?.anchor?{anchor:thread.anchor}:input.anchor?{anchor:input.anchor}:{}),id:input.threadId,resolved:false,updated:now,messages:[...(thread?.messages??[]),{id:input.operationId,author:actor.id,authorName:actor.name,body:input.body!,created:now,...(input.mentions?{mentions:input.mentions}:{})}]});
   // Who should hear about this message: @-mentioned members, and for replies everyone already in the thread.
   const notify:CommentNotice[]=[];if(input.type==='create'||input.type==='reply'){const seen=new Set([actor.id]),excerpt=input.body!.replace(/\s+/g,' ').trim().slice(0,120);
    for(const user of input.mentions??[])if(!seen.has(user)){seen.add(user);notify.push({user,kind:'mention',thread:input.threadId,excerpt});}
    if(thread)for(const message of thread.messages)if(!seen.has(message.author)){seen.add(message.author);notify.push({user:message.author,kind:'reply',thread:input.threadId,excerpt});}}
   receipts.set(key,signature);try{validateOfficeComments(doc);}catch(error){throw new OfficeCommentError(error instanceof Error?error.message:'办公评论超过容量上限',413);}const state=Y.encodeStateAsUpdate(doc);if(state.length>16*1024*1024)throw new OfficeCommentError('办公状态超过 16 MB');
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(state),scope,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(now,workspace,id);return {update:Y.encodeStateAsUpdate(doc,vector),repeated:false,notify};
  }finally{doc.destroy();}
 }).immediate();
}
