import {validateOfficeComments} from '../../../../common/nbstore/src/office-comments';
import * as Y from 'yjs';
import {MAX_OFFICE_TEXT_LENGTH,validTextSettings} from '../../../../common/nbstore/src/office';

export {MAX_OFFICE_TEXT_LENGTH};
export function validateOfficeTextUpdate(current:Uint8Array|null,update:Uint8Array){
  const doc=new Y.Doc();const text=doc.getText('text');
  try{
    if(current)Y.applyUpdate(doc,current);
    Y.applyUpdate(doc,update);
    if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('文本更新缺少依赖，请重新同步');
    validateOfficeComments(doc);if([...doc.share.keys()].some(key=>!['officeComments','officeCommentOperations','text','textSettings'].includes(key))||text._map.size)throw new Error('文本更新包含非法根结构');
    const settings=doc.getMap('textSettings');
    if(settings._start||settings.size&&!validTextSettings(Object.fromEntries(settings.entries())))throw new Error('文本编码或换行设置无效');
    if(text.length>MAX_OFFICE_TEXT_LENGTH)throw new Error('文本超过 100 万字符限制');
    if(text.toDelta().some((part:{insert:unknown;attributes?:unknown})=>typeof part.insert!=='string'||part.attributes))throw new Error('TXT 只允许纯文本内容');
  }catch(error){throw new Error(error instanceof Error?error.message:'文本更新无效');}
  finally{doc.destroy();}
}
