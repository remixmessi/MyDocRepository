import {encodeDelimited} from './office';
export type DelimitedFormat='csv'|'tsv';
export type DelimitedTextMode='escaped'|'raw';
/** Explicit transformation, never a claim that every consuming program is safe. */
export function delimitedText(value:string,mode:DelimitedTextMode){
 if(mode!=='escaped'&&mode!=='raw')throw new Error('表格文本导出方式无效');
 return mode==='escaped'&&(/^[\t\r\n]/.test(value)||/^[\s\u0000-\u001f\u007f]*[=+\-@＝＋－＠]/u.test(value))?"'"+value:value;
}
export function encodeDelimitedText(rows:readonly (readonly string[])[],format:DelimitedFormat,mode:DelimitedTextMode){
 if(format!=='csv'&&format!=='tsv'||mode!=='escaped'&&mode!=='raw')throw new Error('表格文本导出设置无效');
 const width=Math.max(0,...rows.map(row=>row.length));if(rows.length>10000||width*rows.length>200000)throw new Error('导出矩形区域超过 20 万单元格或 1 万行');
 const values=rows.map(row=>Array.from({length:width},(_,c)=>delimitedText(row[c]??'',mode)));
 const bytes=new TextEncoder().encode('\ufeff'+(values.length?encodeDelimited(values,format==='csv'?',':'\t','\r\n'):''));
 if(bytes.length>16*1024*1024)throw new Error('导出文件超过 16 MB');return bytes;
}
