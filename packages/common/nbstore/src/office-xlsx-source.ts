/** `skippedRules`/`skippedCharts` are only present when the import dropped some; older imports never have them. */
export interface XlsxSource {name:string;sheetCount:number;hasStyles:boolean;textConversions:boolean;skippedRules?:number;skippedCharts?:number;}
const count=(value:unknown)=>Number.isSafeInteger(value)&&Number(value)>=1&&Number(value)<=1_000_000;
export function validXlsxSource(value:unknown):value is XlsxSource{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const v=value as Record<string,unknown>,optional=['skippedRules','skippedCharts'].filter(key=>key in v);
 return Object.keys(v).length===4+optional.length&&optional.every(key=>count(v[key]))&&typeof v.name==='string'&&!!v.name.trim()&&v.name.length<=240&&!/[\x00-\x1f]/.test(v.name)&&Number.isSafeInteger(v.sheetCount)&&Number(v.sheetCount)>=1&&Number(v.sheetCount)<=1000&&typeof v.hasStyles==='boolean'&&typeof v.textConversions==='boolean';
}

/** The imported first sheet is the only available formula context. */
export function spreadsheetSheetName(source:unknown):string{return validXlsxSource(source)?source.name:'当前工作表';}
