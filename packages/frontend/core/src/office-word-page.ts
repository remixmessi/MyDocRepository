import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {DEFAULT_WORD_PAGE,validWordPage,wordPageKey,type WordPageLayout} from '../../../common/nbstore/src/office-word-page';
export function getWordPage(doc:Y.Doc):WordPageLayout{
 const value=doc.getMap('pageSettings').get('layout');if(value===undefined)return {...DEFAULT_WORD_PAGE};if(!validWordPage(value))throw new Error('页面设置无效');return {...value};
}
export function setWordPage(doc:Y.Doc,page:WordPageLayout,base?:WordPageLayout){
 if(!validWordPage(page))throw new Error('页面尺寸或页边距无效，请保留至少 18 磅正文区域');
 if(base&&wordPageKey(getWordPage(doc))!==wordPageKey(base))throw new Error('页面设置已被其他成员修改，请重新载入后再保存');
 doc.transact(()=>doc.getMap('pageSettings').set('layout',{...page}),LOCAL_ORIGIN);
}
export function setWordPageBreak(text:Y.Text,value:boolean){
 if(typeof value!=='boolean')throw new Error('段前分页设置无效');if(!text.doc||text.doc.getArray<Y.Text>('paragraphs').toArray().indexOf(text)<0)throw new Error('段落已删除');
 text.doc.transact(()=>text.setAttribute('pageBreakBefore',value),LOCAL_ORIGIN);
}
