import {Inflate} from 'fflate';
import {MAX_OFFICE_FILE_BYTES,decodeOfficeText} from '../../../../common/nbstore/src/office';

export const OFFICE_PACKAGE_LIMITS={entries:2048,totalBytes:64*1024*1024,entryBytes:16*1024*1024,xmlBytes:8*1024*1024,xmlDepth:128,xmlElements:200000};
type Entry={name:string;flags:number;method:number;crc:number;size:number;compressed:number;offset:number;dataStart:number};
const invalid=():never=>{throw new Error('办公 ZIP 包结构无效或已损坏');};
const xmlPart=(name:string)=>/\.(xml|rels)$/i.test(name);
function safeName(bytes:Uint8Array,utf8:boolean){
 if(!utf8&&bytes.some(byte=>byte>127))throw new Error('ZIP 文件名编码不受支持');
 const name=new TextDecoder('utf-8',{fatal:true}).decode(bytes);
 if(!name||name.length>1024||name.startsWith('/')||/[\\:\u0000-\u001f\u007f]/.test(name))throw new Error('ZIP 条目路径无效');
 const parts=name.replace(/\/$/,'').split('/');
 if(parts.some(part=>!part||part==='.'||part==='..'))throw new Error('ZIP 条目路径无效');
 return name;
}
/** Inspect every central/local header before allocating expanded data. Nothing is extracted to disk. */
export function readOfficePackage(bytes:Uint8Array):Record<string,Uint8Array>{
 if(bytes.length>MAX_OFFICE_FILE_BYTES)throw new Error('原件超过 20 MB');
 const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
 const u16=(offset:number)=>{if(offset<0||offset+2>bytes.length)return invalid();return view.getUint16(offset,true);};
 const u32=(offset:number)=>{if(offset<0||offset+4>bytes.length)return invalid();return view.getUint32(offset,true);};
 let end=-1;
 for(let offset=bytes.length-22;offset>=Math.max(0,bytes.length-65557);offset--){if(u32(offset)===0x06054b50&&offset+22+u16(offset+20)===bytes.length){end=offset;break;}}
 if(end<0)return invalid();
 const count=u16(end+10),directorySize=u32(end+12),directoryStart=u32(end+16);
 if(u16(end+4)||u16(end+6)||u16(end+8)!==count||count===65535||directorySize===0xffffffff||directoryStart===0xffffffff)throw new Error('不支持分卷或 ZIP64 办公文件');
 if(count>OFFICE_PACKAGE_LIMITS.entries)throw new Error('ZIP 条目数量超过 2048');
 if(directoryStart+directorySize!==end)return invalid();
 const entries:Entry[]=[],names=new Set<string>();let cursor=directoryStart,total=0;
 for(let index=0;index<count;index++){
  if(cursor+46>end||u32(cursor)!==0x02014b50)return invalid();
  const flags=u16(cursor+8),method=u16(cursor+10),crc=u32(cursor+16),compressed=u32(cursor+20),size=u32(cursor+24),nameLength=u16(cursor+28),extraLength=u16(cursor+30),commentLength=u16(cursor+32),offset=u32(cursor+42);
  const next=cursor+46+nameLength+extraLength+commentLength;if(next>end||u16(cursor+34))return invalid();
  if(size===0xffffffff||compressed===0xffffffff||offset===0xffffffff)throw new Error('不支持 ZIP64 办公文件');
  if(flags&0x41)throw new Error('不支持加密的办公 ZIP 包');
  if(method!==0&&method!==8)throw new Error('不支持该 ZIP 压缩方法');
  const rawName=bytes.subarray(cursor+46,cursor+46+nameLength),name=safeName(rawName,!!(flags&0x800));
  const normalized=name.toLowerCase();if(names.has(normalized))throw new Error('ZIP 包包含重复路径');names.add(normalized);
  const limit=xmlPart(name)?OFFICE_PACKAGE_LIMITS.xmlBytes:OFFICE_PACKAGE_LIMITS.entryBytes;
  if(size>limit)throw new Error(xmlPart(name)?'XML 部件超过 8 MB':'ZIP 单个条目超过 16 MB');
  total+=size;if(total>OFFICE_PACKAGE_LIMITS.totalBytes)throw new Error('ZIP 展开大小超过 64 MB');
  if(offset+30>directoryStart||u32(offset)!==0x04034b50)return invalid();
  const localNameLength=u16(offset+26),localExtraLength=u16(offset+28),dataStart=offset+30+localNameLength+localExtraLength;
  if(dataStart+compressed>directoryStart||u16(offset+6)!==flags||u16(offset+8)!==method||localNameLength!==nameLength)return invalid();
  if(!rawName.every((byte,i)=>bytes[offset+30+i]===byte))return invalid();
  if(!(flags&8)&&(u32(offset+14)!==crc||u32(offset+18)!==compressed||u32(offset+22)!==size))return invalid();
  if(method===0&&size!==compressed)return invalid();
  entries.push({name,flags,method,crc,size,compressed,offset,dataStart});cursor=next;
 }
 if(cursor!==end)return invalid();
 const ordered=entries.slice().sort((a,b)=>a.offset-b.offset);let previousEnd=0;
 for(const entry of ordered){if(entry.offset<previousEnd)return invalid();previousEnd=entry.dataStart+entry.compressed;}
 const files:Record<string,Uint8Array>=Object.create(null);
 for(const entry of entries){
  const compressed=bytes.subarray(entry.dataStart,entry.dataStart+entry.compressed);
  let data:Uint8Array;
  if(entry.method===0)data=compressed.slice();
  else{
   const chunks:Uint8Array[]=[];let actual=0,finished=false;
   const inflate=new Inflate((chunk,final)=>{
    actual+=chunk.length;
    if(actual>entry.size)throw new Error('ZIP 实际展开大小与声明不符');
    chunks.push(chunk.slice());finished=final;
   });
   // Small compressed chunks bound temporary output even if the directory lies about size.
   if(!compressed.length)inflate.push(compressed,true);
   for(let position=0;position<compressed.length;position+=1024)inflate.push(compressed.subarray(position,position+1024),position+1024>=compressed.length);
   if(!finished||actual!==entry.size)return invalid();
   data=new Uint8Array(actual);let position=0;for(const chunk of chunks){data.set(chunk,position);position+=chunk.length;}
  }
  if(data.length!==entry.size||crc32(data)!==entry.crc)throw new Error('ZIP 条目校验失败');
  if(xmlPart(entry.name))readOfficeXml(data);
  files[entry.name]=data;
 }
 return files;
}
const crcTable=Uint32Array.from({length:256},(_,index)=>{let value=index;for(let bit=0;bit<8;bit++)value=(value&1)?0xedb88320^(value>>>1):value>>>1;return value>>>0;});
function crc32(bytes:Uint8Array){let value=0xffffffff;for(const byte of bytes)value=crcTable[(value^byte)&255]^(value>>>8);return (value^0xffffffff)>>>0;}

/** Strict decoding and structural quotas, not an OOXML schema validator. */
export function readOfficeXml(bytes:Uint8Array){
 if(bytes.length>OFFICE_PACKAGE_LIMITS.xmlBytes)throw new Error('XML 部件超过 8 MB');
 const encoding=bytes[0]===0&&bytes[1]===60?'utf-16be':bytes[0]===60&&bytes[1]===0?'utf-16le':undefined;
 const decoded=decodeOfficeText(bytes,encoding),text=decoded.text;
 if(/<!\s*(DOCTYPE|ENTITY)\b/i.test(text))throw new Error('办公 XML 禁止 DTD 和实体声明');
 const declaration=/^\s*<\?xml\s[^?]*\bencoding\s*=\s*['"]([^'"]+)['"]/i.exec(text)?.[1];
 if(declaration&&!/^utf-?(8|16|16le|16be)$/i.test(declaration))throw new Error('办公 XML 编码不受支持');
 if(declaration){const declared=declaration.toLowerCase().replace('-',''),actual=decoded.encoding.replace('-','');if(declared==='utf16'?!actual.startsWith('utf16'):declared!==actual)throw new Error('XML 声明编码与实际字节不符');}
 const stack:string[]=[];let cursor=0,elements=0,roots=0;
 while(cursor<text.length){
  const start=text.indexOf('<',cursor);
  if(start<0){if(!stack.length&&text.slice(cursor).trim())throw new Error('XML 根元素外存在内容');break;}
  if(!stack.length&&text.slice(cursor,start).trim())throw new Error('XML 根元素外存在内容');
  if(text.startsWith('<!--',start)){const end=text.indexOf('-->',start+4);if(end<0)throw new Error('XML 注释未闭合');cursor=end+3;continue;}
  if(text.startsWith('<![CDATA[',start)){const end=text.indexOf(']]>',start+9);if(end<0||!stack.length)throw new Error('XML CDATA 无效');cursor=end+3;continue;}
  if(text.startsWith('<?',start)){const end=text.indexOf('?>',start+2);if(end<0)throw new Error('XML 声明未闭合');cursor=end+2;continue;}
  let quote='',end=start+1;
  for(;end<text.length;end++){const char=text[end];if(quote){if(char===quote)quote='';}else if(char==='"'||char==="'")quote=char;else if(char==='>')break;}
  if(end===text.length)throw new Error('XML 标签未闭合');
  const tag=text.slice(start+1,end),match=/^(\/?)([A-Za-z_][\w.:-]*)(?:\s|\/|$)/.exec(tag);
  if(!match)throw new Error('XML 标签无效');
  if(match[1]){if(stack.pop()!==match[2])throw new Error('XML 标签不匹配');}
  else{
   if(++elements>OFFICE_PACKAGE_LIMITS.xmlElements)throw new Error('XML 元素数量超过限制');
   if(!stack.length&&++roots>1)throw new Error('XML 存在多个根元素');
   if(stack.length+1>OFFICE_PACKAGE_LIMITS.xmlDepth)throw new Error('XML 嵌套深度超过 128');
   if(!tag.endsWith('/'))stack.push(match[2]);
  }
  cursor=end+1;
 }
 if(stack.length||roots!==1)throw new Error('XML 结构不完整');
 return text;
}
