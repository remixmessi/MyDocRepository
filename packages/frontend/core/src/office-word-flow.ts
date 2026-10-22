import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import type {WordBlock} from '../../../common/nbstore/src/office-word-table';

export function wordTextIdentity(text:Y.Text){const id=text._item?.id;return id?`${id.client}:${id.clock}`:undefined;}

function location(text:Y.Text){
 const doc=text.doc,blocks=doc?.getArray<WordBlock>('paragraphs'),index=blocks?.toArray().indexOf(text)??-1;
 if(!doc||!blocks||index<0)throw new Error('正文位置已改变，请重新选择');
 return {doc,blocks,index};
}
function boundary(text:Y.Text,index:number){
 const value=text.toString();
 if(!Number.isInteger(index)||index<0||index>value.length||index>0&&index<value.length&&/[\uD800-\uDBFF]/.test(value[index-1])&&/[\uDC00-\uDFFF]/.test(value[index]))throw new Error('文字位置无效');
}
function room(blocks:Y.Array<WordBlock>){if(blocks.length>=10000)throw new Error('正文内容数量已达到上限');}

/** The empty writing surface always has a caret; structural controls are unnecessary. */
export function wordWritingEnd(doc:Y.Doc){
 const blocks=doc.getArray<WordBlock>('paragraphs'),last=blocks.length?blocks.get(blocks.length-1):undefined;
 if(last instanceof Y.Text)return last;
 room(blocks);const text=new Y.Text();doc.transact(()=>blocks.push([text]),LOCAL_ORIGIN);return text;
}
export function splitWordParagraph(text:Y.Text,start:number,end=start){
 const {doc,blocks,index}=location(text);boundary(text,start);boundary(text,end);if(end<start)throw new Error('文字选区无效');room(blocks);
 const next=new Y.Text();let offset=0;
 const tail:ReturnType<Y.Text['toDelta']>=[];
 for(const part of text.toDelta()){const value=String(part.insert??''),from=Math.max(0,end-offset);if(from<value.length)tail.push({...part,insert:value.slice(from)});offset+=value.length;}
 next.applyDelta(tail);
 for(const [key,value]of Object.entries(text.getAttributes()))if(key!=='pageBreakBefore'&&(key!=='heading'||end<text.length))next.setAttribute(key,value);
 doc.transact(()=>{text.delete(start,text.length-start);blocks.insert(index+1,[next]);},LOCAL_ORIGIN);
 return next;
}
export function adjacentWordParagraph(text:Y.Text,direction:'backward'|'forward'){
 const {blocks,index}=location(text),position=index+(direction==='backward'?-1:1),next=position>=0&&position<blocks.length?blocks.get(position):undefined;
 return next instanceof Y.Text?next:undefined;
}
/** Merge only adjacent text, leaving tables and images as structural boundaries. */
export function mergeWordParagraph(text:Y.Text,direction:'backward'|'forward'){
 const {doc,blocks,index}=location(text),neighbor=adjacentWordParagraph(text,direction);if(!neighbor)return;
 const target=direction==='backward'?neighbor:text,source=direction==='backward'?text:neighbor,offset=target.length;
 const delta=source.toDelta();
 doc.transact(()=>{target.applyDelta([{retain:offset},...delta]);blocks.delete(direction==='backward'?index:index+1,1);},LOCAL_ORIGIN);
 return {text:target,offset};
}
