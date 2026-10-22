export type OfficeKind='word'|'spreadsheet'|'pdf'|'presentation'|'text';
export type TextEncoding='utf-8'|'utf-16le'|'utf-16be';
export type TextLineEnding='lf'|'crlf'|'cr';
export interface OfficeTextSettings {encoding:TextEncoding;bom:boolean;lineEnding:TextLineEnding;}
export const DEFAULT_TEXT_SETTINGS:OfficeTextSettings={encoding:'utf-8',bom:false,lineEnding:'lf'};
export const MAX_OFFICE_TEXT_LENGTH=1_000_000;
export const MAX_OFFICE_FILE_BYTES=20*1024*1024;
/** Formats accepted by the office import endpoint; TXT has its own wizard. */
export const OFFICE_IMPORT_EXTENSIONS=['docx','docm','dotx','dotm','xlsx','xlsm','xltx','xltm','csv','tsv','pdf','pptx','pptm','potx','potm'];
const extension=(name:string)=>name.includes('.')?name.toLowerCase().split('.').pop()!:'';
export function officeImportable(name:string){return OFFICE_IMPORT_EXTENSIONS.includes(extension(name));}
export function officeDelimited(name:string){return ['csv','tsv'].includes(extension(name));}
export function officeImportProblem(name:string,size:number){
 if(['doc','xls','ppt','dot','xlt','pot'].includes(extension(name)))return '暂不支持旧版 .doc / .xls / .ppt，请在 Office 中另存为 .docx / .xlsx / .pptx 后导入';
 if(!officeImportable(name))return '仅支持 Word、Excel、CSV/TSV、PowerPoint 和 PDF 文件';
 if(size>MAX_OFFICE_FILE_BYTES)return '原件超过 20 MB';
 return undefined;
}
export function validTextSettings(value:unknown):value is OfficeTextSettings{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;
 const data=value as Record<string,unknown>;
 return Object.keys(data).length===3&&typeof data.encoding==='string'&&['utf-8','utf-16le','utf-16be'].includes(data.encoding)&&typeof data.bom==='boolean'&&typeof data.lineEnding==='string'&&['lf','crlf','cr'].includes(data.lineEnding);
}
export type DelimitedOptions={delimiter:','|'\t';encoding:TextEncoding;hasHeader:boolean;quoteFields:boolean;}

export function officeKindFromFilename(name:string):{kind:OfficeKind;mime:string;sourceFormat:string}|undefined{
 const ext=name.toLowerCase().split('.').pop()??'';
 if(['docx','docm','dotx','dotm'].includes(ext))return {kind:'word',mime:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',sourceFormat:ext};
 if(['xlsx','xlsm','xltx','xltm','csv','tsv'].includes(ext))return {kind:'spreadsheet',mime:ext==='csv'?'text/csv':ext==='tsv'?'text/tab-separated-values':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',sourceFormat:ext};
 if(ext==='pdf')return {kind:'pdf',mime:'application/pdf',sourceFormat:ext};
 if(['pptx','pptm','potx','potm'].includes(ext))return {kind:'presentation',mime:'application/vnd.openxmlformats-officedocument.presentationml.presentation',sourceFormat:ext};
 if(['txt','log','text'].includes(ext))return {kind:'text',mime:'text/plain',sourceFormat:ext};
 return undefined;
}

export function decodeOfficeText(bytes:Uint8Array,selected?:TextEncoding):{text:string;encoding:TextEncoding;bom:boolean}{
 if(selected&&!['utf-8','utf-16le','utf-16be'].includes(selected))throw new Error('不支持该文本编码');
 let encoding:TextEncoding=selected??'utf-8',start=0,bom=false;
 if(bytes.length>=2&&bytes[0]===0xff&&bytes[1]===0xfe){encoding='utf-16le';start=2;bom=true;}
 else if(bytes.length>=2&&bytes[0]===0xfe&&bytes[1]===0xff){encoding='utf-16be';start=2;bom=true;}
 else if(bytes.length>=3&&bytes[0]===0xef&&bytes[1]===0xbb&&bytes[2]===0xbf){encoding='utf-8';start=3;bom=true;}
 if(bom&&selected&&selected!==encoding)throw new Error('所选编码与文件 BOM 不符');
 const text=new TextDecoder(encoding,{fatal:true,ignoreBOM:true}).decode(bytes.subarray(start));
 return {text,encoding,bom};
}

export function encodeOfficeText(text:string,encoding:TextEncoding='utf-8',bom=false){
 if(!['utf-8','utf-16le','utf-16be'].includes(encoding))throw new Error('不支持该文本编码');
 for(let index=0;index<text.length;index++){
  const unit=text.charCodeAt(index);
  if(unit>=0xd800&&unit<=0xdbff){const next=text.charCodeAt(++index);if(!(next>=0xdc00&&next<=0xdfff))throw new Error('文本包含未配对的 Unicode 代理字符，无法无损导出');}
  else if(unit>=0xdc00&&unit<=0xdfff)throw new Error('文本包含未配对的 Unicode 代理字符，无法无损导出');
 }
 const value=new TextEncoder().encode(text);
 if(encoding==='utf-8')return bom?withPrefix(value,[0xef,0xbb,0xbf]):value;
 const utf16=new Uint8Array(text.length*2);let offset=0;
 for(const char of text){const code=char.codePointAt(0)!;const units=code>0xffff?[0xd800+((code-0x10000)>>10),0xdc00+((code-0x10000)&0x3ff)]:[code];for(const unit of units){if(encoding==='utf-16le'){utf16[offset++]=unit&255;utf16[offset++]=unit>>8;}else{utf16[offset++]=unit>>8;utf16[offset++]=unit&255;}}}
 const body=utf16.slice(0,offset);return bom?withPrefix(body,encoding==='utf-16le'?[0xff,0xfe]:[0xfe,0xff]):body;
}

function withPrefix(bytes:Uint8Array,prefix:number[]){const result=new Uint8Array(prefix.length+bytes.length);result.set(prefix);result.set(bytes,prefix.length);return result;}
export function inspectOfficeText(bytes:Uint8Array,encoding?:TextEncoding){
 if(bytes.length>MAX_OFFICE_FILE_BYTES)throw new Error('原件超过 20 MB');
 const decoded=decodeOfficeText(bytes,encoding);
 if(decoded.text.length>MAX_OFFICE_TEXT_LENGTH)throw new Error('文本超过 100 万字符限制');
 const endings=new Set<TextLineEnding>();
 for(const match of decoded.text.matchAll(/\r\n|\r|\n/g))endings.add(match[0]==='\r\n'?'crlf':match[0]==='\r'?'cr':'lf');
 return {...decoded,mixed:endings.size>1,lineEnding:endings.values().next().value??'lf' as TextLineEnding,normalized:decoded.text.replace(/\r\n|\r/g,'\n')};
}
export function prepareOfficeTextImport(bytes:Uint8Array,encoding?:TextEncoding,lineEnding?:TextLineEnding){
 if(lineEnding&&!['lf','crlf','cr'].includes(lineEnding))throw new Error('换行设置无效');
 const source=inspectOfficeText(bytes,encoding);
 if(source.mixed&&!lineEnding)throw new Error('文件包含混合换行，请明确选择规范化方式');
 return {text:source.normalized,settings:{encoding:source.encoding,bom:source.bom,lineEnding:lineEnding??source.lineEnding} satisfies OfficeTextSettings};
}
export function exportOfficeText(text:string,settings:OfficeTextSettings){
 if(!validTextSettings(settings))throw new Error('文本导出设置无效');
 const normalized=text.replace(/\r\n|\r/g,'\n');
 return encodeOfficeText(normalized.replaceAll('\n',settings.lineEnding==='crlf'?'\r\n':settings.lineEnding==='cr'?'\r':'\n'),settings.encoding,settings.bom);
}

export function parseDelimited(text:string,delimiter=',' as ','|'\t'){
 const rows:string[][]=[];let row:string[]=[],field='',quoted=false;
 for(let i=0;i<text.length;i++){const ch=text[i];if(quoted){if(ch==='"'&&text[i+1]==='"'){field+='"';i++;}else if(ch==='"')quoted=false;else field+=ch;continue;}if(ch==='"'&&field.length===0){quoted=true;continue;}if(ch===delimiter){row.push(field);field='';continue;}if(ch==='\n'||ch==='\r'){if(ch==='\r'&&text[i+1]==='\n')i++;row.push(field);rows.push(row);row=[];field='';continue;}field+=ch;}
 if(quoted)throw new Error('CSV/TSV 引号未闭合');if(field.length||row.length)rows.push([...row,field]);return rows;
}

export function encodeDelimited(rows:readonly (readonly string[])[],delimiter=',' as ','|'\t',lineEnding='\r\n'){
 const quote=(value:string)=>{const needs=value.includes(delimiter)||/["\r\n]/.test(value);const escaped=value.replaceAll('"','""');return needs?'"'+escaped+'"':escaped;};
 return rows.map(row=>row.map(value=>quote(String(value))).join(delimiter)).join(lineEnding)+lineEnding;
}
