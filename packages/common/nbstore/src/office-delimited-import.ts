import {decodeOfficeText,MAX_OFFICE_FILE_BYTES,type TextEncoding} from './office';
export type DelimitedImportOptions={encoding?:TextEncoding;delimiter:','|'\t';quoteMode:'csv'|'none'};
export type DelimitedImportSettings={encoding:TextEncoding;bom:boolean;delimiter:','|'\t';quoteMode:'csv'|'none'};
export function validDelimitedImportOptions(value:unknown):value is DelimitedImportOptions{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const o=value as Record<string,unknown>;
 return Object.keys(o).every(key=>['encoding','delimiter','quoteMode'].includes(key))&&(o.encoding===undefined||typeof o.encoding==='string'&&['utf-8','utf-16le','utf-16be'].includes(o.encoding))&&typeof o.delimiter==='string'&&[',','\t'].includes(o.delimiter)&&typeof o.quoteMode==='string'&&['csv','none'].includes(o.quoteMode)&&!(o.delimiter===','&&o.quoteMode==='none');
}
export function validDelimitedImportSettings(value:unknown):value is DelimitedImportSettings{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const {bom,...options}=value as Record<string,unknown>;
 return Object.keys(value).length===4&&typeof bom==='boolean'&&options.encoding!==undefined&&validDelimitedImportOptions(options);
}
/** Strict bounded state machine shared by the preview and isolated server parser. */
export function prepareDelimitedImport(bytes:Uint8Array,options:DelimitedImportOptions){
 if(bytes.length>MAX_OFFICE_FILE_BYTES)throw new Error('原件超过 20 MB');if(!validDelimitedImportOptions(options))throw new Error('CSV/TSV 导入设置无效');
 const source=decodeOfficeText(bytes,options.encoding);return {rows:parseDelimitedGrid(source.text,options.delimiter,options.quoteMode),settings:{encoding:source.encoding,bom:source.bom,delimiter:options.delimiter,quoteMode:options.quoteMode} satisfies DelimitedImportSettings};
}
/** Parse clipboard text without decoding away a leading BOM or changing field characters. */
export function parseDelimitedGrid(text:string,delimiter:','|'\t',quoteMode:'csv'|'none'='csv'){
 const rows:string[][]=[];let row:string[]=[],field='',quoted=false,closed=false,touched=false,cells=0;
 const append=(value:string)=>{field+=value;if(field.length>10000)throw new Error('单元格超过 1 万字符');};
 const finishField=()=>{if(++cells>200000)throw new Error('表格单元格超过 20 万');row.push(field);field='';closed=false;};
 const finishRow=()=>{finishField();if(rows.length>=10000)throw new Error('表格行数超过 1 万');rows.push(row);row=[];touched=false;};
 for(let i=0;i<text.length;i++){
  const char=text[i];touched=true;
  if(quoted){if(char==='"'){if(text[i+1]==='"'){append('"');i++;}else{quoted=false;closed=true;}}else append(char);continue;}
  if(char===delimiter){finishField();continue;}
  if(char==='\r'||char==='\n'){if(char==='\r'&&text[i+1]==='\n')i++;finishRow();continue;}
  if(closed)throw new Error('结束引号后只能是分隔符或换行，请检查引号模式');
  if(quoteMode==='csv'&&char==='"'){if(field.length)throw new Error('字段内双引号需要转义，请检查引号模式');quoted=true;continue;}
  append(char);
 }
 if(quoted)throw new Error('CSV/TSV 引号未闭合');if(touched||row.length||field.length||closed)finishRow();
 return rows;
}
