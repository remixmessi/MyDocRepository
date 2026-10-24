import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import {trackedReplace,wordVisibleRanges,wordVisibleText,type WordAuthor} from '../../../common/nbstore/src/office-word-revision';
import {normalizeWordNoteRefs,WORD_NOTE_CHAR} from '../../../common/nbstore/src/office-word-reference';
/** Literal find/replace across body paragraphs and table cells, like Word's 查找和替换 without wildcards. */
export type WordMatch={text:Y.Text;start:number;end:number};
export const WORD_REPLACE_LIMIT=5000;
export function findWordMatches(doc:Y.Doc,query:string,matchCase=false):WordMatch[]{
 if(!query)return [];const needle=matchCase?query:query.toLowerCase(),hits:WordMatch[]=[];
 // Search the text as it reads; matches inside tracked deletions are skipped.
 for(const text of wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray())){const full=text.toString();for(const range of wordVisibleRanges(text)){const value=matchCase?full.slice(range.start,range.end):full.slice(range.start,range.end).toLowerCase();let from=0,index:number;while((index=value.indexOf(needle,from))>=0){hits.push({text,start:range.start+index,end:range.start+index+query.length});from=index+query.length;if(hits.length>WORD_REPLACE_LIMIT)return hits;}}}
 return hits;
}
/** Replacement keeps the character formatting of the first replaced character, as Word does. */
function replaceAt(match:WordMatch,replacement:string){
 let position=0,attributes:Record<string,unknown>|undefined;for(const part of match.text.toDelta()){const length=typeof part.insert==='string'?part.insert.length:1;if(match.start<position+length){attributes=part.attributes;break;}position+=length;}
 match.text.delete(match.start,match.end-match.start);if(replacement)match.text.insert(match.start,replacement,{...attributes,note:null});
}
/** With an author (track changes on), the match is marked deleted and the replacement inserted as a revision. */
export function replaceWordMatch(doc:Y.Doc,match:WordMatch,replacement:string,author?:WordAuthor){
 if(!match.text.doc)throw new Error('匹配的文字已删除，请重新查找');doc.transact(()=>{if(author)trackedReplace(match.text,match.start,match.end,replacement,author);else replaceAt(match,replacement);normalizeWordNoteRefs(match.text);},LOCAL_ORIGIN);
}
export function replaceAllWord(doc:Y.Doc,query:string,replacement:string,matchCase=false,author?:WordAuthor){
 const hits=findWordMatches(doc,query,matchCase);if(hits.length>WORD_REPLACE_LIMIT)throw new Error('匹配超过 5000 处，请缩小查找范围');
 // Replace from the end of each paragraph so earlier offsets stay valid.
 doc.transact(()=>{for(const match of [...hits].reverse())if(author)trackedReplace(match.text,match.start,match.end,replacement,author);else replaceAt(match,replacement);for(const text of new Set(hits.map(hit=>hit.text)))normalizeWordNoteRefs(text);},LOCAL_ORIGIN);return hits.length;
}
/** Headings for the navigation pane (body paragraphs only, as in Word's 导航窗格). */
export function wordOutline(doc:Y.Doc){
 return doc.getArray<WordBlock>('paragraphs').toArray().flatMap(block=>block instanceof Y.Text&&Number(block.getAttribute('heading'))>0?[{text:block,level:Number(block.getAttribute('heading')),title:wordVisibleText(block).replaceAll(WORD_NOTE_CHAR,'').replace(/\s+/g,' ').trim()||'（空标题）'}]:[]);
}
/** Word-style counts: each CJK character is a word; runs of other letters/digits count as one word. */
export function wordCounts(doc:Y.Doc){
 // Tracked deletions do not count, as in Word's 字数统计 with changes shown.
 const texts=wordBlockTexts(doc.getArray<WordBlock>('paragraphs').toArray()),all=texts.map(text=>wordVisibleText(text).replaceAll(WORD_NOTE_CHAR,'')).join('\n'),scripts=/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/gu;
 const chars=Array.from(all.replace(/[\n\r\u2029]/g,'')),cjk=(all.match(scripts)??[]).length,latin=(all.replace(scripts,' ').match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu)??[]).length;
 return {words:cjk+latin,characters:chars.length,charactersNoSpaces:chars.filter(char=>!/\s/.test(char)).length,paragraphs:texts.filter(text=>wordVisibleText(text).trim()).length};
}
