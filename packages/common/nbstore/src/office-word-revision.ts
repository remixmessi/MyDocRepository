import {Y} from '../../crdt/src/index';
import {wordBlockTexts,isWordTable,type WordBlock} from './office-word-table';
import {validWordFontSize,validWordColor,validWordHighlight,validWordVertAlign,validWordFontFamily} from './office-word-font';
import {validWordAlignment,validWordLeftIndent,validWordRightIndent,validWordFirstLineIndent,validWordLineSpacing,validWordParagraphSpace} from './office-word-layout';
import {validWordList} from './office-word-list';

/**
 * Word track changes. While tracking is on:
 * - typed text carries an `ins` mark and deleted text stays in place with a `del` mark (struck through);
 * - character formatting records the previous values in a `fmt` run mark, paragraph formatting in a `pfmt`
 *   paragraph attribute (Word's rPrChange / pPrChange);
 * - splitting a paragraph marks the new paragraph `insBreak`, joining two marks the second `delBreak` (Word's
 *   inserted/deleted paragraph mark), so the paragraphs stay apart until the change is accepted;
 * - inserted and deleted pictures carry `ins`/`del` on the picture block.
 * Every mark holds the revision id, author and time. Table structure changes are not tracked.
 */
export type WordRevisionMark={id:string;author:string;name:string;date:number};
export type WordAuthor={id:string;name:string};
export const WORD_REVISION_KEYS=['ins','del'] as const;
export const validWordRevisionMark=(value:unknown):value is WordRevisionMark=>!!value&&typeof value==='object'&&!Array.isArray(value)&&Object.keys(value).length===4
 &&typeof (value as WordRevisionMark).id==='string'&&/^[a-zA-Z0-9-]{1,80}$/.test((value as WordRevisionMark).id)&&typeof (value as WordRevisionMark).author==='string'&&/^[a-zA-Z0-9-]{1,100}$/.test((value as WordRevisionMark).author)
 &&typeof (value as WordRevisionMark).name==='string'&&(value as WordRevisionMark).name.length>=1&&(value as WordRevisionMark).name.length<=160&&Number.isSafeInteger((value as WordRevisionMark).date)&&(value as WordRevisionMark).date>0;
export const wordTracking=(doc:Y.Doc)=>doc.share.has('pageSettings')&&doc.getMap('pageSettings').get('trackChanges')===true;
type Delta={insert:string;attributes?:Record<string,unknown>};
const deltaOf=(text:Y.Text)=>text.toDelta() as Delta[];
/** Attributes at each character (UTF-16 index), so ranges can be split by their marks. */
function marksAt(text:Y.Text){const out:(Record<string,unknown>|undefined)[]=[];for(const part of deltaOf(text))for(let i=0;i<String(part.insert).length;i++)out.push(part.attributes);return out;}
const GROUP_MS=60_000;
/** Reuse a neighbouring revision of the same author made within a minute, so continuous typing is one change. */
function neighbour(marks:(Record<string,unknown>|undefined)[],index:number,kind:'ins'|'del',author:WordAuthor,now:number){
 const mark=marks[index]?.[kind];return validWordRevisionMark(mark)&&mark.author===author.id&&now-mark.date<GROUP_MS?mark:undefined;
}
/**
 * Apply a plain-text edit (`before` → `after`, as typed in the editor) with tracking. Deleted characters get a `del`
 * mark instead of being removed — except the author's own pending insertions, which simply disappear, as in Word.
 * Returns the caret offset; `direction` tells where a pure deletion leaves it.
 */
export function applyTrackedEdit(text:Y.Text,before:string,after:string,author:WordAuthor,origin:unknown,direction:'backward'|'forward'='backward',now=Date.now()){
 let start=0,end=before.length,nextEnd=after.length;
 while(start<end&&start<nextEnd&&before[start]===after[start])start++;
 while(end>start&&nextEnd>start&&before[end-1]===after[nextEnd-1]){end--;nextEnd--;}
 if(start===end&&start===nextEnd)return start;
 if(start>0&&/[\uD800-\uDBFF]/.test(before[start-1]))start--;
 if(end<before.length&&/[\uDC00-\uDFFF]/.test(before[end])){end++;nextEnd++;}
 if(before!==text.toString())throw new Error('正文已被其他成员修改，请重新输入');
 const marks=marksAt(text),inserted=after.slice(start,nextEnd);
 const del=neighbour(marks,direction==='forward'?start-1:end,'del',author,now)??{id:crypto.randomUUID(),author:author.id,name:author.name,date:now};
 const ins=neighbour(marks,start-1,'ins',author,now)??{id:crypto.randomUUID(),author:author.id,name:author.name,date:now};
 let removedBefore=0;
 text.doc!.transact(()=>{
  // Right to left so earlier indexes stay valid while own insertions are removed.
  for(let i=end;i>start;){
   const at=i-1,own=validWordRevisionMark(marks[at]?.ins)&&(marks[at]!.ins as WordRevisionMark).author===author.id,gone=validWordRevisionMark(marks[at]?.del);let j=at;
   while(j>start&&(validWordRevisionMark(marks[j-1]?.ins)&&(marks[j-1]!.ins as WordRevisionMark).author===author.id)===own&&validWordRevisionMark(marks[j-1]?.del)===gone)j--;
   if(gone){/* already deleted */}else if(own){text.delete(j,i-j);removedBefore+=i-j;}else text.format(j,i-j,{del:del});
   i=j;
  }
  if(inserted){const at=start+(direction==='forward'?0:end-start-removedBefore),base=marks[start-1]??{};const attributes:Record<string,unknown>={...base,ins};delete attributes.del;text.insert(at,inserted,attributes);}
 },origin);
 // Typing ends after the new text; Backspace leaves the caret before the struck text, Delete after it.
 const kept=end-start-removedBefore;
 return inserted?start+kept+inserted.length:direction==='forward'?start+kept:start;
}
export type WordRevisionKind='ins'|'del'|'fmt'|'pfmt'|'insBreak'|'delBreak'|'insImage'|'delImage';
/**
 * One revision run. Text kinds (ins/del/fmt) cover [start,end) of `text`; paragraph kinds (pfmt, insBreak, delBreak)
 * have `text` = the paragraph and start = end = 0; picture kinds have `block` and no text. `order` sorts in document order.
 */
export type WordRevision={id:string;kind:WordRevisionKind;author:string;name:string;date:number;text?:Y.Text;block?:Y.Map<unknown>;start:number;end:number;excerpt:string;order:number;old?:Record<string,unknown>};
const TEXT_KINDS=['ins','del','fmt'] as const;
/** Document-order key of a caret, matching `WordRevision.order`. */
export function wordRevisionOrder(doc:Y.Doc){
 const keys=new Map<object,number>();let slot=0;
 for(const block of doc.getArray<WordBlock>('paragraphs').toArray()){if(isWordTable(block)){for(const text of wordBlockTexts([block]))keys.set(text,++slot);}else keys.set(block as object,++slot);}
 return (item:object,offset=0)=>(keys.get(item)??0)*1e6+offset;
}
/** All revisions in document order (paragraphs, pictures and table cells); one entry per contiguous run of a revision. */
export function wordRevisions(doc:Y.Doc):WordRevision[]{
 const out:WordRevision[]=[],order=wordRevisionOrder(doc),blocks=doc.getArray<WordBlock>('paragraphs').toArray();
 const mark=(value:unknown)=>validWordRevisionMark(value)?value:undefined;
 const runs=(text:Y.Text)=>{let at=0;
  for(const part of deltaOf(text)){const value=String(part.insert),length=value.length;
   for(const kind of TEXT_KINDS){const raw=part.attributes?.[kind],found=kind==='fmt'?(validWordRunFormatMark(raw)?raw:undefined):mark(raw);if(!found)continue;const last=out[out.length-1],old=kind==='fmt'?(found as WordFormatMark).old:undefined;
    if(last&&last.text===text&&last.kind===kind&&last.id===found.id&&last.end===at&&same(last.old,old)){last.end+=length;last.excerpt=(last.excerpt+value).slice(0,80);}
    else out.push({id:found.id,kind,author:found.author,name:found.name,date:found.date,text,start:at,end:at+length,excerpt:value.slice(0,80),order:order(text,at),...(old?{old}:{})});}
   at+=length;}
 };
 for(const block of blocks){
  if(block instanceof Y.Text){
   const paragraph=(kind:'pfmt'|'insBreak'|'delBreak')=>{const raw=block.getAttribute(kind),found=kind==='pfmt'?(validWordParagraphFormatMark(raw)?raw:undefined):mark(raw);if(found)out.push({id:found.id,kind,author:found.author,name:found.name,date:found.date,text:block,start:0,end:0,excerpt:wordVisibleText(block).slice(0,80),order:order(block),...(kind==='pfmt'?{old:(found as WordFormatMark).old}:{})});};
   paragraph('insBreak');paragraph('delBreak');paragraph('pfmt');runs(block);
  }else if(isWordTable(block)){for(const text of wordBlockTexts([block]))runs(text);}
  else{const image=block as Y.Map<unknown>;if(image.get('kind')!=='image')continue;for(const kind of ['ins','del'] as const){const found=mark(image.get(kind));if(found)out.push({id:found.id,kind:kind==='ins'?'insImage':'delImage',author:found.author,name:found.name,date:found.date,block:image,start:0,end:0,excerpt:String(image.get('alt')).slice(0,80),order:order(image)});}}
 }
 return out.sort((a,b)=>a.order-b.order);
}
/** Join `text` onto the previous text paragraph (keeping the previous paragraph's properties); false if there is none. */
function mergeIntoPrevious(text:Y.Text){
 const blocks=text.doc!.getArray<WordBlock>('paragraphs'),index=blocks.toArray().indexOf(text),previous=index>0?blocks.get(index-1):undefined;
 if(!(previous instanceof Y.Text))return false;previous.applyDelta([{retain:previous.length},...text.toDelta()]);blocks.delete(index,1);return true;
}
/**
 * Accept or reject revisions (one id, or all). Accepting keeps insertions, removes deletions and keeps new formatting;
 * rejecting does the reverse. Text runs are resolved first, then paragraph marks and pictures from the end backwards,
 * so earlier positions stay valid.
 */
export function resolveWordRevisions(doc:Y.Doc,accept:boolean,origin:unknown,id?:string){
 const targets=wordRevisions(doc).filter(revision=>id===undefined||revision.id===id);if(!targets.length)return 0;
 doc.transact(()=>{
  for(const revision of [...targets].reverse()){
   const text=revision.text!,length=revision.end-revision.start;
   if(revision.kind==='ins'||revision.kind==='del'){const remove=revision.kind==='ins'?!accept:accept;if(remove)text.delete(revision.start,length);else text.format(revision.start,length,{[revision.kind]:null});}
   else if(revision.kind==='fmt')text.format(revision.start,length,accept?{fmt:null}:{...revision.old,fmt:null});
   else if(revision.kind==='pfmt'){if(!accept)for(const [key,value]of Object.entries(revision.old??{})){if(value===null)text.removeAttribute(key);else text.setAttribute(key,value);}text.removeAttribute('pfmt');}
  }
  for(const revision of [...targets].reverse()){
   const {text,block}=revision;
   if(revision.kind==='insBreak'){text!.removeAttribute('insBreak');if(!accept){text!.removeAttribute('delBreak');mergeIntoPrevious(text!);}}
   else if(revision.kind==='delBreak'){text!.removeAttribute('delBreak');if(accept)mergeIntoPrevious(text!);}
   else if(revision.kind==='insImage'||revision.kind==='delImage'){
    const key=revision.kind==='insImage'?'ins':'del',remove=revision.kind==='insImage'?!accept:accept,blocks=doc.getArray<WordBlock>('paragraphs'),index=blocks.toArray().indexOf(block as WordBlock);
    if(index<0)continue;if(remove)blocks.delete(index,1);else block!.delete(key);
   }
  }
 },origin);
 return new Set(targets.map(revision=>revision.id)).size;
}
/**
 * Tracked replacement of one range: the range is marked deleted and `value` is inserted after it with the format of
 * the range's first character. Returns the offset just after the inserted text.
 */
export function trackedReplace(text:Y.Text,start:number,end:number,value:string,author:WordAuthor,now=Date.now()){
 const marks=marksAt(text),base:Record<string,unknown>={...(marks[start]??marks[start-1]??{})};delete base.del;delete base.ins;
 const del={id:crypto.randomUUID(),author:author.id,name:author.name,date:now};let offset=end;
 for(let i=end;i>start;){const at=i-1,own=validWordRevisionMark(marks[at]?.ins)&&(marks[at]!.ins as WordRevisionMark).author===author.id,gone=validWordRevisionMark(marks[at]?.del);let j=at;
  while(j>start&&(validWordRevisionMark(marks[j-1]?.ins)&&(marks[j-1]!.ins as WordRevisionMark).author===author.id)===own&&validWordRevisionMark(marks[j-1]?.del)===gone)j--;
  if(own){text.delete(j,i-j);offset-=i-j;}else if(!gone)text.format(j,i-j,{del});i=j;}
 if(value)text.insert(offset,value,{...base,ins:{id:crypto.randomUUID(),author:author.id,name:author.name,date:now}});
 return offset+value.length;
}
/** Character ranges of `text` that are not deleted, mapped back to Y.Text offsets. */
export function wordVisibleRanges(text:Y.Text){const out:{start:number;end:number}[]=[];let at=0;for(const part of deltaOf(text)){const length=String(part.insert).length;if(!validWordRevisionMark(part.attributes?.del)){const last=out[out.length-1];if(last&&last.end===at)last.end+=length;else out.push({start:at,end:at+length});}at+=length;}return out;}
/** Text as it reads with all changes accepted (no deleted runs), for counts and search. */
export function wordVisibleText(text:Y.Text){return deltaOf(text).filter(part=>!validWordRevisionMark(part.attributes?.del)).map(part=>String(part.insert)).join('');}
/** Stable colour per author, as Word gives each reviewer a colour. */
export const WORD_REVISION_COLORS=['B42318','175CD3','027A48','B54708','6941C6','C11574','0E7090','4E5BA6'];
export function wordRevisionColor(author:string){let hash=0;for(const char of author)hash=(hash*31+char.charCodeAt(0))>>>0;return '#'+WORD_REVISION_COLORS[hash%WORD_REVISION_COLORS.length];}

// ── Formatting, paragraph-mark and picture revisions ──
/** Character properties a `fmt` revision can record, with their validators (null in `old` means "was not set"). */
export const WORD_RUN_FORMAT_VALID:Record<string,(value:unknown)=>boolean>={bold:v=>typeof v==='boolean',italic:v=>typeof v==='boolean',underline:v=>typeof v==='boolean',strike:v=>typeof v==='boolean',fontSize:validWordFontSize,color:validWordColor,highlight:validWordHighlight,vertAlign:validWordVertAlign,fontFamily:validWordFontFamily};
/** Paragraph properties a `pfmt` revision can record. */
export const WORD_PARAGRAPH_FORMAT_VALID:Record<string,(value:unknown)=>boolean>={heading:v=>Number.isInteger(v)&&Number(v)>=0&&Number(v)<=3,align:validWordAlignment,leftIndent:validWordLeftIndent,rightIndent:validWordRightIndent,firstLineIndent:validWordFirstLineIndent,list:validWordList,lineSpacing:validWordLineSpacing,spaceBefore:validWordParagraphSpace,spaceAfter:validWordParagraphSpace,pageBreakBefore:v=>typeof v==='boolean'};
export type WordFormatMark=WordRevisionMark&{old:Record<string,unknown>};
const validFormat=(value:unknown,valid:Record<string,(value:unknown)=>boolean>):value is WordFormatMark=>{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const {old,...mark}=value as WordFormatMark;
 return validWordRevisionMark(mark)&&!!old&&typeof old==='object'&&!Array.isArray(old)&&Object.keys(old).length>=1&&Object.entries(old).every(([key,item])=>Object.hasOwn(valid,key)&&(item===null||valid[key](item)));
};
export const validWordRunFormatMark=(value:unknown):value is WordFormatMark=>validFormat(value,WORD_RUN_FORMAT_VALID);
export const validWordParagraphFormatMark=(value:unknown):value is WordFormatMark=>validFormat(value,WORD_PARAGRAPH_FORMAT_VALID);
/** Paragraph attributes that hold revisions (only on body paragraphs, not table cells). */
export const WORD_PARAGRAPH_REVISION_KEYS=['pfmt','insBreak','delBreak'] as const;

/**
 * The author whose edits are tracked, per document; set by the editor while 修订 is on. Formatting helpers consult
 * it so every formatting path (ribbon, shortcuts, format painter, cross-paragraph selection) is tracked alike.
 */
const trackingAuthors=new WeakMap<Y.Doc,WordAuthor>();
export function setWordTrackingAuthor(doc:Y.Doc,author:WordAuthor|undefined){if(author)trackingAuthors.set(doc,author);else trackingAuthors.delete(doc);}
export const wordTrackingAuthor=(doc:Y.Doc|null|undefined)=>doc?trackingAuthors.get(doc):undefined;
const newMark=(author:WordAuthor,now:number):WordRevisionMark=>({id:crypto.randomUUID(),author:author.id,name:author.name,date:now});
const same=(a:unknown,b:unknown)=>JSON.stringify(a??null)===JSON.stringify(b??null);

/**
 * Character formatting with tracking: applies `changes` (null clears) to [start,end) and records, per run, the values
 * before the first tracked change. Text the author inserted themselves and deleted text are formatted untracked, as
 * in Word. A change back to the original values drops the record.
 */
export function formatWordTracked(text:Y.Text,start:number,end:number,changes:Record<string,unknown>,author:WordAuthor,now=Date.now()){
 if(end<=start)return;const marks=marksAt(text),fresh=newMark(author,now);
 const plan:(Record<string,unknown>|null|undefined)[]=[];
 for(let i=start;i<end;i++){
  const current=marks[i]??{},own=validWordRevisionMark(current.ins)&&current.ins.author===author.id;
  if(own||validWordRevisionMark(current.del)){plan.push(undefined);continue;}
  const prior=validWordRunFormatMark(current.fmt)?current.fmt:undefined,old:Record<string,unknown>={...prior?.old};
  for(const [key,value]of Object.entries(changes))if(!Object.hasOwn(old,key)&&!same(current[key],value))old[key]=current[key]??null;
  // Keys whose new value equals the recorded original are no longer a change.
  for(const key of Object.keys(old)){const after=Object.hasOwn(changes,key)?changes[key]:current[key];if(same(after,old[key]))delete old[key];}
  // The author's earlier format revision continues; anything else becomes part of this edit's one new revision.
  const identity=prior&&prior.author===author.id?{id:prior.id,author:prior.author,name:prior.name,date:prior.date}:fresh;
  plan.push(Object.keys(old).length?{...identity,old}:null);
 }
 text.format(start,end-start,changes);
 for(let i=0;i<plan.length;){let j=i+1;while(j<plan.length&&same(plan[j],plan[i]))j++;if(plan[i]!==undefined)text.format(start+i,j-i,{fmt:plan[i]});i=j;}
}
/** Paragraph attribute with tracking; `undefined` removes it. Records the value before the first tracked change. */
export function setWordParagraphAttributeTracked(text:Y.Text,key:string,value:unknown,author:WordAuthor|undefined,now=Date.now()){
 const before=text.getAttribute(key);
 if(value===undefined)text.removeAttribute(key);else text.setAttribute(key,value);
 if(!author||!Object.hasOwn(WORD_PARAGRAPH_FORMAT_VALID,key)||validWordRevisionMark(text.getAttribute('insBreak'))&&(text.getAttribute('insBreak') as WordRevisionMark).author===author.id)return;
 const prior=text.getAttribute('pfmt'),mark=validWordParagraphFormatMark(prior)?prior:undefined,old:Record<string,unknown>={...mark?.old};
 if(!Object.hasOwn(old,key)){if(same(before,value))return;old[key]=before??null;}else if(same(old[key],value))delete old[key];
 if(!Object.keys(old).length){text.removeAttribute('pfmt');return;}
 text.setAttribute('pfmt',{...(mark?{id:mark.id,author:mark.author,name:mark.name,date:mark.date}:newMark(author,now)),old});
}
/** A paragraph split while tracking: the new paragraph's leading break is an insertion. */
export function markWordParagraphInserted(next:Y.Text,author:WordAuthor,now=Date.now()){next.setAttribute('insBreak',newMark(author,now));next.removeAttribute('pfmt');next.removeAttribute('delBreak');}
/**
 * Join `second` onto the previous paragraph while tracking. Undoing the author's own split merges for real; other
 * breaks are marked `delBreak` and stay until accepted. Returns true when the paragraphs were actually merged.
 */
export function deleteWordParagraphBreak(second:Y.Text,author:WordAuthor,now=Date.now()){
 const own=second.getAttribute('insBreak');if(validWordRevisionMark(own)&&own.author===author.id)return true;
 if(!validWordRevisionMark(second.getAttribute('delBreak')))second.setAttribute('delBreak',newMark(author,now));return false;
}
/** Picture insertion/deletion while tracking; deleting the author's own pending picture removes it. Returns true if removed. */
export function markWordImageInserted(image:Y.Map<unknown>,author:WordAuthor,now=Date.now()){image.set('ins',newMark(author,now));}
export function deleteWordImageTracked(image:Y.Map<unknown>,author:WordAuthor,now=Date.now()){
 const own=image.get('ins');if(validWordRevisionMark(own)&&own.author===author.id)return true;
 if(!validWordRevisionMark(image.get('del')))image.set('del',newMark(author,now));return false;
}
/** Every revision mark in the document (runs, paragraph marks, paragraph formatting, pictures), for validation. */
export function* wordRevisionMarks(doc:Y.Doc):Generator<WordRevisionMark>{
 const strip=(value:WordRevisionMark)=>({id:value.id,author:value.author,name:value.name,date:value.date});
 for(const block of doc.getArray<WordBlock>('paragraphs').toArray()){
  if(block instanceof Y.Map&&block.get('kind')==='image'){for(const key of ['ins','del']){const value=block.get(key);if(validWordRevisionMark(value))yield value;}continue;}
  if(block instanceof Y.Text){for(const key of ['insBreak','delBreak']){const value=block.getAttribute(key);if(validWordRevisionMark(value))yield value;}const format=block.getAttribute('pfmt');if(validWordParagraphFormatMark(format))yield strip(format);}
  for(const text of wordBlockTexts([block]))for(const part of deltaOf(text)){for(const key of WORD_REVISION_KEYS){const value=part.attributes?.[key];if(validWordRevisionMark(value))yield value;}const format=part.attributes?.fmt;if(validWordRunFormatMark(format))yield strip(format);}
 }
}
