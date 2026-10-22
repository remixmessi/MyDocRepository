import {Y} from '../../crdt/src/index';
import {wordBlockTexts,type WordBlock} from './office-word-table';

/**
 * Word track changes. While tracking is on, typed text carries an `ins` mark and deleted text stays in place with a
 * `del` mark (shown struck through) until someone accepts or rejects it. Both marks hold the revision id, author and
 * time. Paragraph splits/merges, formatting and table/image changes are not tracked.
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
export type WordRevision={id:string;kind:'ins'|'del';author:string;name:string;date:number;text:Y.Text;start:number;end:number;excerpt:string};
/** All revisions in document order (paragraphs and table cells); one entry per contiguous run of a revision. */
export function wordRevisions(doc:Y.Doc):WordRevision[]{
 const out:WordRevision[]=[];
 for(const text of wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray())){let at=0;
  for(const part of deltaOf(text)){const value=String(part.insert),length=value.length;
   for(const kind of WORD_REVISION_KEYS){const mark=part.attributes?.[kind];if(!validWordRevisionMark(mark))continue;const last=out[out.length-1];
    if(last&&last.text===text&&last.kind===kind&&last.id===mark.id&&last.end===at){last.end+=length;last.excerpt=(last.excerpt+value).slice(0,80);}
    else out.push({id:mark.id,kind,author:mark.author,name:mark.name,date:mark.date,text,start:at,end:at+length,excerpt:value.slice(0,80)});}
   at+=length;}
 }
 return out;
}
/** Accept or reject revisions (one id, or all). Accepting keeps insertions and removes deletions; rejecting the reverse. */
export function resolveWordRevisions(doc:Y.Doc,accept:boolean,origin:unknown,id?:string){
 const targets=wordRevisions(doc).filter(revision=>id===undefined||revision.id===id);if(!targets.length)return 0;
 doc.transact(()=>{
  for(const revision of [...targets].reverse()){
   const remove=revision.kind==='ins'?!accept:accept;
   if(remove)revision.text.delete(revision.start,revision.end-revision.start);else revision.text.format(revision.start,revision.end-revision.start,{[revision.kind]:null});
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
