/**
 * Word headers and footers: left/centre/right text with {页码}/{总页数} fields, optionally different on the first
 * page and on even pages. Each slot holds up to five lines separated by "\n".
 * Text is stored as `pageSettings.headerFooter`; pictures as `pageSettings.headerImages` (a Y.Map of Word image maps
 * keyed `scope.part.slot`, see `wordHeaderImageKey`), so they share the document's media handling with body images.
 * Exported as header/footer parts with PAGE and NUMPAGES fields.
 */
export type WordHeaderLine={left:string;center:string;right:string};
export type WordHeaderPair={header:WordHeaderLine;footer:WordHeaderLine};
export type WordHeaderFooter=WordHeaderPair&{first?:WordHeaderPair;even?:WordHeaderPair};
export type WordHeaderScope='main'|'first'|'even';
export type WordHeaderPart='header'|'footer';
export type WordHeaderSlot=keyof WordHeaderLine;
export const WORD_PAGE_TOKEN='{页码}',WORD_PAGES_TOKEN='{总页数}';
export const EMPTY_WORD_LINE:WordHeaderLine={left:'',center:'',right:''};
export const WORD_HEADER_SLOTS=['left','center','right'] as const;
export const WORD_HEADER_SCOPES=['main','first','even'] as const;
export const WORD_HEADER_LINES=5;
// Up to five lines; no other control characters or Unicode line/paragraph separators (U+2028, U+2029).
const INVALID=new RegExp('[\\x00-\\x09\\x0b-\\x1f\\u2028\\u2029]');
const record=(value:unknown):value is Record<string,unknown>=>!!value&&typeof value==='object'&&!Array.isArray(value);
const keys=(value:Record<string,unknown>,required:string[],optional:string[]=[])=>required.every(key=>key in value)&&Object.keys(value).every(key=>required.includes(key)||optional.includes(key));
const validText=(text:unknown)=>typeof text==='string'&&text.length<=255&&!INVALID.test(text)&&text.split('\n').length<=WORD_HEADER_LINES;
export function validWordHeaderLine(value:unknown):value is WordHeaderLine{
 return record(value)&&keys(value,[...WORD_HEADER_SLOTS])&&WORD_HEADER_SLOTS.every(slot=>validText(value[slot]));
}
const validPair=(value:unknown):value is WordHeaderPair=>record(value)&&keys(value,['header','footer'])&&validWordHeaderLine(value.header)&&validWordHeaderLine(value.footer);
export function validWordHeaderFooter(value:unknown):value is WordHeaderFooter{
 return record(value)&&keys(value,['header','footer'],['first','even'])&&validWordHeaderLine(value.header)&&validWordHeaderLine(value.footer)
  &&(value.first===undefined||validPair(value.first))&&(value.even===undefined||validPair(value.even));
}
export const emptyWordLine=(line:WordHeaderLine)=>WORD_HEADER_SLOTS.every(slot=>!line[slot]);
/** Which variant a page (1-based) uses: the first-page one on page 1, the even one on even pages, else the default. */
export function wordHeaderScope(value:Pick<WordHeaderFooter,'first'|'even'>|undefined,page:number):WordHeaderScope{
 return page===1&&value?.first?'first':page%2===0&&value?.even?'even':'main';
}
export const wordHeaderPair=(value:WordHeaderFooter,scope:WordHeaderScope):WordHeaderPair|undefined=>scope==='main'?value:value[scope];
/** Lines for a page (1-based) with fields replaced by numbers. */
export function wordHeaderLines(value:WordHeaderFooter|undefined,page:number,pages:number){
 if(!value)return;const source=wordHeaderPair(value,wordHeaderScope(value,page))!,fill=(text:string)=>text.replaceAll(WORD_PAGE_TOKEN,String(page)).replaceAll(WORD_PAGES_TOKEN,String(pages));
 const line=(item:WordHeaderLine):WordHeaderLine=>({left:fill(item.left),center:fill(item.center),right:fill(item.right)});
 return {header:line(source.header),footer:line(source.footer)};
}
/** Text split into literal parts and page fields, for export (PAGE/NUMPAGES fields) and print (CSS counters). */
export function wordHeaderParts(text:string):({text:string}|{field:'PAGE'|'NUMPAGES'})[]{
 return text.split(/(\{页码\}|\{总页数\})/).filter(Boolean).map(part=>part===WORD_PAGE_TOKEN?{field:'PAGE' as const}:part===WORD_PAGES_TOKEN?{field:'NUMPAGES' as const}:{text:part});
}
/** One picture per slot, shown above that slot's text. */
export const wordHeaderImageKey=(scope:WordHeaderScope,part:WordHeaderPart,slot:WordHeaderSlot)=>`${scope}.${part}.${slot}`;
export function parseWordHeaderImageKey(key:string){
 const [scope,part,slot,...rest]=key.split('.');
 return !rest.length&&(WORD_HEADER_SCOPES as readonly string[]).includes(scope)&&['header','footer'].includes(part)&&(WORD_HEADER_SLOTS as readonly string[]).includes(slot)?{scope:scope as WordHeaderScope,part:part as WordHeaderPart,slot:slot as WordHeaderSlot}:undefined;
}
/** Header pictures are kept small: at most 600 × 200 pixels. */
export const WORD_HEADER_IMAGE_MAX={width:600,height:200};
