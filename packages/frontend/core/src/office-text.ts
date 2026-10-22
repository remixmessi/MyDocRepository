import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {DEFAULT_TEXT_SETTINGS,validTextSettings,type OfficeTextSettings} from '../../../common/nbstore/src/office';

export function officeTextSettings(doc:Y.Doc):OfficeTextSettings{
  const settings=doc.getMap('textSettings');
  if(!settings.size)return {...DEFAULT_TEXT_SETTINGS};
  const value=Object.fromEntries(settings.entries());
  if(!validTextSettings(value))throw new Error('文本编码或换行设置无效');
  return value;
}
export function setOfficeTextSettings(doc:Y.Doc,change:Partial<OfficeTextSettings>){
  const next={...officeTextSettings(doc),...change};if(!validTextSettings(next))throw new Error('文本导出设置无效');
  const settings=doc.getMap('textSettings');doc.transact(()=>{for(const [key,value]of Object.entries(next))if(settings.get(key)!==value)settings.set(key,value);},LOCAL_ORIGIN);
}

// Resolve edit boundaries against the shared document, including after IME composition.
export function applyOfficeTextEdit(text:Y.Text,before:string,after:string,reference=text){
  let start=0,end=before.length,nextEnd=after.length;
  while(start<end&&start<nextEnd&&before[start]===after[start])start++;
  while(end>start&&nextEnd>start&&before[end-1]===after[nextEnd-1]){end--;nextEnd--;}
  if(start===end&&start===nextEnd)return;
  if(start>0&&/[\uD800-\uDBFF]/.test(before[start-1]))start--;
  if(end<before.length&&/[\uDC00-\uDFFF]/.test(before[end])){end++;nextEnd++;}
  const left=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromTypeIndex(reference,start),text.doc!);
  const right=Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromTypeIndex(reference,end),text.doc!);
  if(!left||!right)throw new Error('文本位置已失效，请保留草稿后重试');
  text.doc!.transact(()=>{if(right.index>left.index)text.delete(left.index,right.index-left.index);text.insert(left.index,after.slice(start,nextEnd));},LOCAL_ORIGIN);
}
