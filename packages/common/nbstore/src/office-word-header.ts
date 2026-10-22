/**
 * Word headers and footers: left/centre/right text with {页码}/{总页数} fields, optionally different on the first page.
 * Stored as `pageSettings.headerFooter`; exported as header/footer parts with PAGE and NUMPAGES fields.
 */
export type WordHeaderLine={left:string;center:string;right:string};
export type WordHeaderFooter={header:WordHeaderLine;footer:WordHeaderLine;first?:{header:WordHeaderLine;footer:WordHeaderLine}};
export const WORD_PAGE_TOKEN='{页码}',WORD_PAGES_TOKEN='{总页数}';
export const EMPTY_WORD_LINE:WordHeaderLine={left:'',center:'',right:''};
const SLOTS=['left','center','right'] as const;
// One line of text: no control characters or line/paragraph separators (U+2028, U+2029).
const INVALID=new RegExp('[\\x00-\\x1f\\u2028\\u2029]');
const record=(value:unknown):value is Record<string,unknown>=>!!value&&typeof value==='object'&&!Array.isArray(value);
const keys=(value:Record<string,unknown>,required:string[],optional:string[]=[])=>required.every(key=>key in value)&&Object.keys(value).every(key=>required.includes(key)||optional.includes(key));
export function validWordHeaderLine(value:unknown):value is WordHeaderLine{
 return record(value)&&keys(value,[...SLOTS])&&SLOTS.every(slot=>typeof value[slot]==='string'&&(value[slot] as string).length<=255&&!INVALID.test(value[slot] as string));
}
export function validWordHeaderFooter(value:unknown):value is WordHeaderFooter{
 return record(value)&&keys(value,['header','footer'],['first'])&&validWordHeaderLine(value.header)&&validWordHeaderLine(value.footer)
  &&(value.first===undefined||record(value.first)&&keys(value.first,['header','footer'])&&validWordHeaderLine(value.first.header)&&validWordHeaderLine(value.first.footer));
}
export const emptyWordLine=(line:WordHeaderLine)=>SLOTS.every(slot=>!line[slot]);
/** Lines for a page (1-based) with fields replaced by numbers. */
export function wordHeaderLines(value:WordHeaderFooter|undefined,page:number,pages:number){
 if(!value)return;const source=page===1&&value.first?value.first:value,fill=(text:string)=>text.replaceAll(WORD_PAGE_TOKEN,String(page)).replaceAll(WORD_PAGES_TOKEN,String(pages));
 const line=(item:WordHeaderLine):WordHeaderLine=>({left:fill(item.left),center:fill(item.center),right:fill(item.right)});
 return {header:line(source.header),footer:line(source.footer)};
}
/** Text split into literal parts and page fields, for export (PAGE/NUMPAGES fields) and print (CSS counters). */
export function wordHeaderParts(text:string):({text:string}|{field:'PAGE'|'NUMPAGES'})[]{
 return text.split(/(\{页码\}|\{总页数\})/).filter(Boolean).map(part=>part===WORD_PAGE_TOKEN?{field:'PAGE' as const}:part===WORD_PAGES_TOKEN?{field:'NUMPAGES' as const}:{text:part});
}
