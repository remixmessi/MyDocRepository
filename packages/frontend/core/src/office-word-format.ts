import {assertWordTextLive} from './office-word-location';
import {validWordFontSize,validWordColor,validWordHighlight,validWordFontFamily,validWordVertAlign} from '../../../common/nbstore/src/office-word-font';
import {validWordAlignment,validWordLeftIndent,validWordRightIndent,validWordFirstLineIndent,validWordLineSpacing,validWordParagraphSpace,type WordAlignment} from '../../../common/nbstore/src/office-word-layout';
import {validWordLink} from '../../../common/nbstore/src/office-word-link';
import {validWordList,type WordList} from '../../../common/nbstore/src/office-word-list';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
export type WordMark='bold'|'italic'|'underline'|'strike';
/** Character properties set from a value list (`undefined` clears back to the paragraph default). */
export type WordFontKey='fontSize'|'color'|'highlight'|'fontFamily'|'vertAlign';
const FONT_VALID:Record<WordFontKey,(value:unknown)=>boolean>={fontSize:validWordFontSize,color:validWordColor,highlight:validWordHighlight,fontFamily:validWordFontFamily,vertAlign:validWordVertAlign};
/** Direct character formatting that “清除格式” removes; links are content, not formatting, and stay. */
export const WORD_RUN_FORMATS=['bold','italic','underline','strike','fontSize','color','highlight','fontFamily','vertAlign'] as const;
export type WordSelection={start:Y.RelativePosition;end:Y.RelativePosition};
export function captureWordSelection(text:Y.Text,start:number,end:number):WordSelection|undefined{
 if(start===end)return;
 // Keep formatting boundaries outside surrogate pairs.
 const value=text.toString();
 if(start>0&&/[\uDC00-\uDFFF]/.test(value[start]??''))start--;
 if(end<value.length&&/[\uDC00-\uDFFF]/.test(value[end]??''))end++;
 return {start:Y.createRelativePositionFromTypeIndex(text,start,0),end:Y.createRelativePositionFromTypeIndex(text,end,-1)};
}
export function resolveWordSelection(text:Y.Text,selection:WordSelection){
 if(!text.doc)return;
 const start=Y.createAbsolutePositionFromRelativePosition(selection.start,text.doc),end=Y.createAbsolutePositionFromRelativePosition(selection.end,text.doc);
 if(!start||!end||start.type!==text||end.type!==text||end.index<=start.index)return;
 return {start:start.index,end:end.index};
}
export function wordMarkActive(text:Y.Text,start:number,end:number,mark:WordMark){
 let position=0;
 for(const part of text.toDelta()){
  const next=position+(typeof part.insert==='string'?part.insert.length:1);
  if(next>start&&position<end&&part.attributes?.[mark]!==true)return false;
  position=next;
 }
 return end>start;
}
export function toggleWordMark(text:Y.Text,selection:WordSelection,mark:WordMark){
 assertWordTextLive(text);
 const range=resolveWordSelection(text,selection);if(!range)throw new Error('所选文字已删除，请重新选择');
 const value=!wordMarkActive(text,range.start,range.end,mark);
 text.doc!.transact(()=>text.format(range.start,range.end-range.start,{[mark]:value}),LOCAL_ORIGIN);
 return range;
}

export function setWordHeading(text:Y.Text,heading:number){
 if(!Number.isInteger(heading)||heading<0||heading>3)throw new Error('标题级别无效');
 assertWordTextLive(text);
 text.doc!.transact(()=>text.setAttribute('heading',heading),LOCAL_ORIGIN);
}

export function setWordList(text:Y.Text,kind:'none'|'bullet'|'decimal'){
 if(!['none','bullet','decimal'].includes(kind))throw new Error('列表类型无效');
 assertWordTextLive(text);
 const doc=text.doc!,paragraphs=doc.getArray('paragraphs').toArray(),index=paragraphs.indexOf(text);
 const prior=index>0?paragraphs[index-1]:undefined;
 const current=text.getAttribute('list'),previous=prior instanceof Y.Text?prior.getAttribute('list'):undefined;
 const list:WordList|undefined=kind==='none'?undefined:validWordList(current)&&current.kind===kind?current:validWordList(previous)&&previous.kind===kind?{...previous}:{kind,id:crypto.randomUUID(),start:1};
 doc.transact(()=>{if(list)text.setAttribute('list',list);else text.removeAttribute('list');},LOCAL_ORIGIN);
}
export function wordListMarkers(paragraphs:Y.Text[]){
 const counters=new Map<string,number>();
 return paragraphs.map(paragraph=>{const list=paragraph.getAttribute('list');if(!validWordList(list))return '';const number=counters.get(list.id)??list.start;counters.set(list.id,number+1);return list.kind==='bullet'?'•':`${number}.`;});
}

export function setWordLink(text:Y.Text,selection:WordSelection,href?:string){
 if(href!==undefined&&!validWordLink(href))throw new Error('请输入不含账号密码的 HTTP/HTTPS 网页地址');
 assertWordTextLive(text);
 const range=resolveWordSelection(text,selection);if(!range)throw new Error('所选文字已删除，请重新选择');
 text.doc!.transact(()=>text.format(range.start,range.end-range.start,{href:href??null}),LOCAL_ORIGIN);
 return range;
}
export function wordSelectionLink(text:Y.Text,start:number,end:number){
 let position=0,href:string|undefined,first=true;
 for(const part of text.toDelta()){
  const next=position+(typeof part.insert==='string'?part.insert.length:1);
  if(next>start&&position<end){if(!validWordLink(part.attributes?.href))return '';if(!first&&href!==part.attributes.href)return '';href=part.attributes.href;first=false;}
  position=next;
 }
 return href??'';
}

export function setWordAlignment(text:Y.Text,align:WordAlignment){
 if(!validWordAlignment(align))throw new Error('段落对齐方式无效');
 assertWordTextLive(text);
 text.doc!.transact(()=>text.setAttribute('align',align),LOCAL_ORIGIN);
}
export function setWordLeftIndent(text:Y.Text,leftIndent:number){
 if(!validWordLeftIndent(leftIndent))throw new Error('左缩进须在 0–144 磅范围内');
 assertWordTextLive(text);
 text.doc!.transact(()=>text.setAttribute('leftIndent',leftIndent),LOCAL_ORIGIN);
}

export function setWordRightIndent(text:Y.Text,rightIndent:number){
 if(!validWordRightIndent(rightIndent))throw new Error('右缩进须在 0–144 磅范围内');
 assertWordTextLive(text);
 text.doc!.transact(()=>text.setAttribute('rightIndent',rightIndent),LOCAL_ORIGIN);
}
export function setWordFirstLineIndent(text:Y.Text,firstLineIndent:number){
 if(!validWordFirstLineIndent(firstLineIndent))throw new Error('首行或悬挂缩进须在 0–144 磅范围内');
 assertWordTextLive(text);
 text.doc!.transact(()=>text.setAttribute('firstLineIndent',firstLineIndent),LOCAL_ORIGIN);
}

export function setWordSpacing(text:Y.Text,key:'lineSpacing'|'spaceBefore'|'spaceAfter',value:number|undefined){
 if(value!==undefined&&!(key==='lineSpacing'?validWordLineSpacing(value):validWordParagraphSpace(value)))throw new Error(key==='lineSpacing'?'行距须在 0.5–5 倍之间':'段前/段后间距须在 0–72 磅之间');
 assertWordTextLive(text);
 text.doc!.transact(()=>{if(value===undefined)text.removeAttribute(key);else text.setAttribute(key,value);},LOCAL_ORIGIN);
}
export function clearWordFormatting(text:Y.Text,selection:WordSelection){
 assertWordTextLive(text);const range=resolveWordSelection(text,selection);if(!range)throw new Error('所选文字已删除，请重新选择');
 text.doc!.transact(()=>text.format(range.start,range.end-range.start,Object.fromEntries(WORD_RUN_FORMATS.map(key=>[key,null]))),LOCAL_ORIGIN);return range;
}
/** Character formatting at a position, for the format painter. */
export function wordFormatAt(text:Y.Text,index:number){
 let position=0;for(const part of text.toDelta()){const next=position+(typeof part.insert==='string'?part.insert.length:1);if(index<next||next===text.length&&index===next){const marks=part.attributes??{};return Object.fromEntries(WORD_RUN_FORMATS.filter(key=>marks[key]!==undefined).map(key=>[key,marks[key]]));}position=next;}
 return {} as Record<string,unknown>;
}
export function applyWordFormat(text:Y.Text,selection:WordSelection,format:Record<string,unknown>){
 assertWordTextLive(text);const range=resolveWordSelection(text,selection);if(!range)throw new Error('所选文字已删除，请重新选择');
 text.doc!.transact(()=>text.format(range.start,range.end-range.start,Object.fromEntries(WORD_RUN_FORMATS.map(key=>[key,format[key]??null]))),LOCAL_ORIGIN);return range;
}
export function wordSelectionFont(text:Y.Text,start:number,end:number,key:WordFontKey):number|string|undefined{
 let position=0,first=true,value:number|string|undefined;
 for(const part of text.toDelta()){
  const next=position+(typeof part.insert==='string'?part.insert.length:1);
  if(next>start&&position<end){const current=part.attributes?.[key];if(!first&&current!==value)return 'mixed';value=current;first=false;}
  position=next;
 }
 return value;
}
export function setWordFont(text:Y.Text,selection:WordSelection,key:WordFontKey,value:number|string|undefined){
 if(!Object.hasOwn(FONT_VALID,key))throw new Error('文字属性无效');
 if(value!==undefined&&!FONT_VALID[key](value))throw new Error(key==='fontSize'?'字号须在 1–200 磅且为半磅步进':key==='color'?'文字颜色无效':key==='highlight'?'突出显示颜色无效':key==='fontFamily'?'字体名称无效':'上下标设置无效');
 assertWordTextLive(text);
 const range=resolveWordSelection(text,selection);if(!range)throw new Error('所选文字已删除，请重新选择');
 const normalized=key==='color'&&typeof value==='string'&&value!=='auto'?value.toUpperCase():value;
 text.doc!.transact(()=>text.format(range.start,range.end-range.start,{[key]:normalized??null}),LOCAL_ORIGIN);return range;
}
