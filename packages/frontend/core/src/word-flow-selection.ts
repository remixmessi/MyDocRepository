import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import type {WordBlock} from '../../../common/nbstore/src/office-word-table';
import {captureWordSelection,setWordFont,formatWordRun,wordMarkActive,clearWordFormatting,type WordMark,type WordFontKey} from './office-word-format';
import {splitWordParagraph} from './office-word-flow';
import {trackedReplace,deleteWordParagraphBreak,type WordAuthor} from '../../../common/nbstore/src/office-word-revision';
export interface WordFlowSelection {start:Y.RelativePosition;end:Y.RelativePosition;}
export function captureWordFlowSelection(first:Y.Text,start:number,last:Y.Text,end:number):WordFlowSelection{
 if(first.doc!==last.doc||!first.doc)throw new Error('文字选区无效');
 for(const [text,index]of [[first,start],[last,end]]as const){const value=text.toString();if(!Number.isInteger(index)||index<0||index>value.length||index>0&&index<value.length&&/[\uD800-\uDBFF]/.test(value[index-1])&&/[\uDC00-\uDFFF]/.test(value[index]))throw new Error('文字位置无效');}
 return {start:Y.createRelativePositionFromTypeIndex(first,start,0),end:Y.createRelativePositionFromTypeIndex(last,end,-1)};
}
export function resolveWordFlowSelection(doc:Y.Doc,selection:WordFlowSelection){
 const a=Y.createAbsolutePositionFromRelativePosition(selection.start,doc),b=Y.createAbsolutePositionFromRelativePosition(selection.end,doc),blocks=doc.getArray<WordBlock>('paragraphs');
 if(!a||!b||!(a.type instanceof Y.Text)||!(b.type instanceof Y.Text))throw new Error('选区所在的正文已删除，请重新选择');
 const all=blocks.toArray(),first=all.indexOf(a.type),last=all.indexOf(b.type);
 if(first<0||last<first||first===last&&b.index<a.index)throw new Error('选区所在的正文已改变，请重新选择');
 const texts=all.slice(first,last+1);if(texts.some(text=>!(text instanceof Y.Text)))throw new Error('此选区跨越表格或图片，请分别编辑正文与对象');
 const ranges=(texts as Y.Text[]).map((text,index)=>({text,start:index===0?a.index:0,end:index===texts.length-1?b.index:text.length}));
 return {blocks,first,last,ranges};
}
export function wordFlowPlainText(doc:Y.Doc,selection:WordFlowSelection){return resolveWordFlowSelection(doc,selection).ranges.map(({text,start,end})=>text.toString().slice(start,end)).join('\n');}
export function formatWordFlow(doc:Y.Doc,selection:WordFlowSelection,key:WordMark|WordFontKey|'clear',value?:number|string){
 const {ranges}=resolveWordFlowSelection(doc,selection),mark=['bold','italic','underline','strike'].includes(key),enabled=mark&&!ranges.filter(r=>r.end>r.start).every(r=>wordMarkActive(r.text,r.start,r.end,key as WordMark));
 doc.transact(()=>{for(const {text,start,end}of ranges){if(end<=start)continue;const selection=captureWordSelection(text,start,end)!;if(key==='clear')clearWordFormatting(text,selection);else if(mark)formatWordRun(text,start,end,{[key]:enabled});else setWordFont(text,selection,key as WordFontKey,value);}},LOCAL_ORIGIN);
}
/**
 * Replace one contiguous text range as one local undo operation, retaining tail marks. With an author (track
 * changes on) every part of the range and the paragraph marks inside it are marked deleted, and the new text is
 * inserted in the first paragraph; typed line breaks become spaces.
 */
export function replaceWordFlow(doc:Y.Doc,selection:WordFlowSelection,value:string,author?:WordAuthor){
 const {blocks,first,last,ranges}=resolveWordFlowSelection(doc,selection),head=ranges[0],tail=ranges[ranges.length-1];
 if(author){let offset=head.start;doc.transact(()=>{for(const range of [...ranges].reverse()){const inserted=range===head?value.replace(/\r\n|\r|\n/g,' '):'';const end=trackedReplace(range.text,range.start,range.end,inserted,author);if(range===head)offset=end;
   // The paragraph marks inside the selection are deleted too; the author's own pending splits are undone.
   if(range!==head&&deleteWordParagraphBreak(range.text,author)){const index=blocks.toArray().indexOf(range.text),previous=blocks.get(index-1);if(previous instanceof Y.Text){range.text.removeAttribute('insBreak');previous.applyDelta([{retain:previous.length},...range.text.toDelta()]);blocks.delete(index,1);}}}},LOCAL_ORIGIN);return {text:head.text,offset};}
 const normalized=value.replace(/\r\n|\r/g,'\n'),lines=normalized.split('\n');if(blocks.length-(last-first)+lines.length-1>10000||normalized.length>1000000)throw new Error('正文内容数量或字符数量超限');
 const suffix:ReturnType<Y.Text['toDelta']>=[];let offset=0;
 for(const part of tail.text.toDelta()){const text=String(part.insert??''),from=Math.max(0,tail.end-offset);if(from<text.length)suffix.push({...part,insert:text.slice(from)});offset+=text.length;}
 let target=head.text,caret=head.start;
 doc.transact(()=>{
  head.text.delete(head.start,head.text.length-head.start);if(last>first)blocks.delete(first+1,last-first);
  if(lines[0])target.insert(caret,lines[0]);caret+=lines[0].length;
  for(const line of lines.slice(1)){target=splitWordParagraph(target,caret);if(line)target.insert(0,line);caret=line.length;}
  target.applyDelta([{retain:target.length},...suffix]);
 },LOCAL_ORIGIN);return {text:target,offset:caret};
}
