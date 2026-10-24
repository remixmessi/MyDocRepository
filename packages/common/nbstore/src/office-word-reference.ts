import {Y} from '../../crdt/src/index';
/**
 * Word 引用与水印: text watermarks (pageSettings.watermark), table-of-contents blocks in `paragraphs`, and
 * footnotes/endnotes. A note reference is one private-use character (WORD_NOTE_CHAR) carrying a `note` id in the
 * paragraph's Y.Text, so offsets, selections, revisions and comments treat it as one character; the note's text lives
 * in the `wordNotes` root (id → {kind, text: Y.Text}).
 */

// ---- watermark --------------------------------------------------------------------------------------------------
export type WordWatermark={text:string;color:string;layout:'diagonal'|'horizontal';size?:number};
export const WORD_WATERMARK_PRESETS=['机密','绝密','严禁复制','草稿','样本','内部资料','CONFIDENTIAL','DRAFT'];
export const DEFAULT_WORD_WATERMARK_COLOR='#C0C0C0';
export function validWordWatermark(value:unknown):value is WordWatermark{
 if(!value||typeof value!=='object'||Array.isArray(value)||value instanceof Y.AbstractType)return false;const v=value as Record<string,unknown>;
 return Object.keys(v).every(key=>['text','color','layout','size'].includes(key))&&typeof v.text==='string'&&v.text.trim()!==''&&v.text===v.text.trim()&&[...v.text].length<=50&&!/[\x00-\x1f\u2028\u2029]/.test(v.text)
  &&typeof v.color==='string'&&/^#[0-9A-F]{6}$/.test(v.color)&&(v.layout==='diagonal'||v.layout==='horizontal')&&(v.size===undefined||Number.isInteger(v.size)&&Number(v.size)>=20&&Number(v.size)<=200);
}
/** Font size in points: a set size, or one that lets the text span about 80% of the body width (diagonal: the diagonal). */
export function wordWatermarkSize(watermark:WordWatermark,page:{width:number;height:number;left:number;right:number;top:number;bottom:number}){
 if(watermark.size)return watermark.size;const width=(page.width-page.left-page.right)/20,height=(page.height-page.top-page.bottom)/20,span=(watermark.layout==='diagonal'?Math.hypot(width,height)*.8:width*.8);
 const ems=[...watermark.text].reduce((total,char)=>total+(/[\u2E80-\uFFFF]/.test(char)?1:.6),0);return Math.round(Math.min(144,Math.max(36,span/Math.max(1,ems))));
}

// ---- table of contents ------------------------------------------------------------------------------------------
export type WordTocLevels=1|2|3;
export const WORD_TOC_LIMIT=20;
export function isWordToc(value:unknown):value is Y.Map<unknown>{return value instanceof Y.Map&&value.get('kind')==='toc';}
export function createWordToc(levels:WordTocLevels=3){const block=new Y.Map<unknown>();block.set('kind','toc');block.set('levels',levels);return block;}
export function validateWordToc(block:Y.Map<unknown>){if(block._start||block.size!==2||block.get('kind')!=='toc'||![1,2,3].includes(block.get('levels') as number))throw new Error('Word 目录结构无效');}
export const wordTocLevels=(block:Y.Map<unknown>)=>(Number(block.get('levels'))||3) as WordTocLevels;

// ---- footnotes and endnotes -------------------------------------------------------------------------------------
export const WORD_NOTE_CHAR='\uE000';
export type WordNoteKind='footnote'|'endnote';
export const WORD_NOTE_LIMITS={notes:1000,text:5000};
export const validWordNoteId=(value:unknown):value is string=>typeof value==='string'&&/^[a-zA-Z0-9-]{1,80}$/.test(value);
export function wordNotes(doc:Y.Doc){return doc.share.has('wordNotes')?doc.getMap<Y.Map<unknown>>('wordNotes'):undefined;}
export function createWordNote(doc:Y.Doc,id:string,kind:WordNoteKind,text=''){
 if(!validWordNoteId(id))throw new Error('注释编号无效');const notes=doc.getMap<Y.Map<unknown>>('wordNotes');if(notes.size>=WORD_NOTE_LIMITS.notes)throw new Error('脚注与尾注最多 1000 条');
 const note=new Y.Map<unknown>(),body=new Y.Text();body.insert(0,text.slice(0,WORD_NOTE_LIMITS.text).replaceAll(WORD_NOTE_CHAR,''));note.set('kind',kind);note.set('text',body);notes.set(id,note);return note;
}
export function validateWordNotes(doc:Y.Doc){
 const notes=wordNotes(doc);if(!notes)return;if(notes._start||notes.size>WORD_NOTE_LIMITS.notes)throw new Error('Word 脚注尾注过多或结构无效');
 for(const [id,note]of notes.entries()){
  if(!validWordNoteId(id)||!(note instanceof Y.Map)||note._start||note.size!==2||!['footnote','endnote'].includes(note.get('kind') as string))throw new Error('Word 脚注尾注结构无效');
  const text=note.get('text');if(!(text instanceof Y.Text)||text.length>WORD_NOTE_LIMITS.text||Object.keys(text.getAttributes()).length||text.toString().includes(WORD_NOTE_CHAR)||/[\x00-\x08\x0b\x0c\x0e-\x1f\u2029]/.test(text.toString())||text.toDelta().some((part:{insert?:unknown;attributes?:object})=>typeof part.insert!=='string'||part.attributes&&Object.keys(part.attributes).length))throw new Error('Word 脚注尾注内容无效');
 }
}
/** Lower-case roman numerals for endnotes (Word's default i, ii, iii …). */
export function wordRoman(value:number){let n=value,out='';for(const [v,s]of [[1000,'m'],[900,'cm'],[500,'d'],[400,'cd'],[100,'c'],[90,'xc'],[50,'l'],[40,'xl'],[10,'x'],[9,'ix'],[5,'v'],[4,'iv'],[1,'i']] as const)while(n>=v){out+=s;n-=v;}return out;}
export type WordNoteRef={id:string;kind:WordNoteKind;number:number;label:string;text:Y.Text;offset:number;deleted:boolean};
/**
 * Note references in reading order (body paragraphs, then cells of each table row by row). Footnotes and endnotes are
 * numbered separately; a reference whose note no longer exists is skipped. The first reference of an id numbers it.
 */
export function wordNoteRefs(doc:Y.Doc,texts:Y.Text[]){
 const notes=wordNotes(doc),refs:WordNoteRef[]=[],numbers=new Map<string,{number:number;label:string}>(),count={footnote:0,endnote:0};if(!notes)return refs;
 for(const text of texts){let offset=0;for(const part of text.toDelta() as {insert:string;attributes?:Record<string,unknown>}[]){
  const value=String(part.insert),id=part.attributes?.note;
  if(value===WORD_NOTE_CHAR&&validWordNoteId(id)){const note=notes.get(id),kind=note?.get('kind') as WordNoteKind|undefined;if(note&&kind){let entry=numbers.get(id);if(!entry){const number=++count[kind];entry={number,label:kind==='footnote'?String(number):wordRoman(number)};numbers.set(id,entry);}refs.push({id,kind,...entry,text,offset,deleted:!!part.attributes?.del});}}
  offset+=value.length;
 }}
 return refs;
}
/** Remove `note` from characters other than the reference character (Y.Text copies formatting onto typed text). */
export function normalizeWordNoteRefs(text:Y.Text){
 // Inherited text merges into the reference's own delta part, so clear character runs around each reference character.
 let offset=0;const fixes:[number,number][]=[];for(const part of text.toDelta() as {insert:string;attributes?:Record<string,unknown>}[]){const value=String(part.insert);if(part.attributes?.note!==undefined&&value!==WORD_NOTE_CHAR){let start=-1;for(let i=0;i<=value.length;i++){const other=i<value.length&&value[i]!==WORD_NOTE_CHAR;if(other&&start<0)start=i;if(!other&&start>=0){fixes.push([offset+start,i-start]);start=-1;}}}offset+=value.length;}
 for(const [at,length]of fixes)text.format(at,length,{note:null});return fixes.length>0;
}
