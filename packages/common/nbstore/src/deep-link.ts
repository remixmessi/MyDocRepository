export interface DocumentLocation {workspace:string;doc:string;mode:'page'|'board';frame?:string;block?:string;}
const validId=(value:string)=>value.length>0&&value.length<=500&&!/[\s\x00-\x1f\x7f/?#&\\]/.test(value);
export function parseDocumentLink(input:string):DocumentLocation|null {
 try{
  if(input.length>2048)return null;const url=new URL(input);
  if(url.protocol!=='mydoc:'||url.hostname!=='document'||url.pathname!==''&&url.pathname!=='/'||url.username||url.password||url.port||url.hash)return null;
  const p=url.searchParams;if([...p.keys()].some(key=>!['workspace','doc','mode','frame','block'].includes(key)||p.getAll(key).length!==1))return null;
  const workspace=p.get('workspace')??'',doc=p.get('doc')??'',frame=p.get('frame')??undefined,block=p.get('block')??undefined,mode=p.get('mode')??(frame?'board':'page');
  if(!validId(workspace)||!validId(doc)||frame!==undefined&&!validId(frame)||block!==undefined&&!validId(block)||!['page','board'].includes(mode)||frame&&mode!=='board'||block&&mode!=='page'||frame&&block)return null;
  return {workspace,doc,mode:mode as 'page'|'board',...(frame?{frame}:{}),...(block?{block}:{})};
 }catch{return null;}
}
export function documentLink(target:DocumentLocation){const params=new URLSearchParams({workspace:target.workspace,doc:target.doc,mode:target.mode});if(target.frame)params.set('frame',target.frame);if(target.block)params.set('block',target.block);const url=`mydoc://document?${params}`;if(!parseDocumentLink(url))throw new Error('文档链接目标无效');return url;}
export function shareDocumentLink(target:DocumentLocation){if(typeof window!=='undefined'&&window.mydocDesktop)return documentLink(target);const url=new URL(location.href);url.hash=new URLSearchParams({workspace:target.workspace,doc:target.doc,...(target.frame?{frame:target.frame}:{}),...(target.block?{block:target.block}:{})}).toString();return url.href;}
