export interface XlsxSource {name:string;sheetCount:number;hasStyles:boolean;textConversions:boolean;}
export function validXlsxSource(value:unknown):value is XlsxSource{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const v=value as Record<string,unknown>;
 return Object.keys(v).length===4&&typeof v.name==='string'&&!!v.name.trim()&&v.name.length<=240&&!/[\x00-\x1f]/.test(v.name)&&Number.isSafeInteger(v.sheetCount)&&Number(v.sheetCount)>=1&&Number(v.sheetCount)<=1000&&typeof v.hasStyles==='boolean'&&typeof v.textConversions==='boolean';
}

/** The imported first sheet is the only available formula context. */
export function spreadsheetSheetName(source:unknown):string{return validXlsxSource(source)?source.name:'当前工作表';}
