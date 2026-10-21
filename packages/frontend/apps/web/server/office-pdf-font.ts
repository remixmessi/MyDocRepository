import {PDFDocument,StandardFonts,type PDFFont} from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';

/**
 * Unicode appearance font for PDF export (form values, flattened notes). Runs inside the restricted parser process,
 * which has no file access: the parent reads the bundled font and prepends it to the job (see `wrapPdfFont`).
 */
const MAGIC=0x31464d50;// "PMF1" little-endian; old packets start with a state length that is never this large.
export const MAX_PDF_FONT_BYTES=16*1024*1024;
export function wrapPdfFont(packet:Uint8Array,font:Uint8Array|undefined){
 if(!font)return packet;if(font.length>MAX_PDF_FONT_BYTES)throw new Error('PDF 字体文件过大');
 const header=Buffer.alloc(8);header.writeUInt32LE(MAGIC,0);header.writeUInt32LE(font.length,4);return Buffer.concat([header,font,packet]);
}
export function unwrapPdfFont(bytes:Uint8Array):{packet:Uint8Array;font?:Uint8Array}{
 const view=Buffer.from(bytes.buffer,bytes.byteOffset,bytes.byteLength);
 if(view.length<8||view.readUInt32LE(0)!==MAGIC)return {packet:bytes};
 const length=view.readUInt32LE(4);if(length>MAX_PDF_FONT_BYTES||length+8>view.length)throw new Error('PDF 字体输入无效');
 return {font:view.subarray(8,8+length),packet:view.subarray(8+length)};
}

/**
 * fontkit's TrueType subsetter writes the short (16-bit) loca table by halving every glyph offset, but copies glyph
 * data without padding; a glyph of odd length then shifts every later glyph and they render blank or garbled.
 * Padding each copied glyph to an even length keeps the halved offsets exact (trailing bytes are ignored by readers).
 */
type Subset={_addGlyph:(gid:number)=>number;glyf:Uint8Array[];offset:number};
const paddedFontkit:typeof fontkit={...fontkit,create(...args:Parameters<typeof fontkit.create>){
 // fontkit's internal subset fields are not in its type definitions.
 const font=fontkit.create(...args),internal=font as unknown as {createSubset:()=>Subset},create=internal.createSubset.bind(font);
 internal.createSubset=()=>{const subset=create(),add=subset._addGlyph;
  subset._addGlyph=function(this:Subset,gid:number){const index=add.call(this,gid),buffer=this.glyf[index];
   if(buffer.length%2){const Kind=buffer.constructor as unknown as {alloc?:(n:number)=>Uint8Array},copy=Kind.alloc?Kind.alloc(buffer.length+1):new Uint8Array(buffer.length+1);copy.set(buffer,0);this.glyf[index]=copy;this.offset+=1;}
   return index;};
  return subset;};
 return font;}};

/** Picks Helvetica when it can encode `text`, otherwise the embedded Unicode font; one embedding per document. */
export class PdfFonts{
 private standard?:Promise<PDFFont>;private unicode?:Promise<PDFFont>;private characters?:Set<number>;
 constructor(private readonly pdf:PDFDocument,private readonly font:Uint8Array|undefined){}
 async for(text:string,purpose:string):Promise<PDFFont>{
  const standard=await (this.standard??=this.pdf.embedFont(StandardFonts.Helvetica));
  try{standard.encodeText(text);return standard;}catch{/* needs the Unicode font */}
  if(!this.font)throw new Error(`${purpose}含有中文等字符，但服务器缺少可嵌入的 Unicode 字体。内容已协作保存，可下载文件备份`);
  const unicode=await (this.unicode??=(this.pdf.registerFontkit(paddedFontkit),this.pdf.embedFont(this.font,{subset:true})));
  this.characters??=new Set(unicode.getCharacterSet());
  const missing=[...new Set([...text].filter(char=>!/\s/.test(char)&&!this.characters!.has(char.codePointAt(0)!)))];
  if(missing.length)throw new Error(`${purpose}含有 PDF 字体无法显示的字符：${missing.slice(0,5).join(' ')}${missing.length>5?' 等':''}（如表情符号），请删除这些字符后导出`);
  return unicode;
 }
}
