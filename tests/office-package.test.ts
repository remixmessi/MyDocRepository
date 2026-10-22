import {describe,it,expect,vi} from 'vitest';
import {zipSync,strToU8,Inflate,Zip,ZipDeflate} from 'fflate';
import {readOfficePackage,readOfficeXml,OFFICE_PACKAGE_LIMITS} from '../packages/frontend/apps/web/server/office-package';
import {encodeOfficeText} from '../packages/common/nbstore/src/office';

function zip(files:Record<string,string>,level:0|6=6){return zipSync(Object.fromEntries(Object.entries(files).map(([name,text])=>[name,strToU8(text)])),{level});}
function entries(bytes:Uint8Array){const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);const end=bytes.length-22,count=view.getUint16(end+10,true);let cursor=view.getUint32(end+16,true);const out:{central:number;local:number}[]=[];for(let i=0;i<count;i++){out.push({central:cursor,local:view.getUint32(cursor+42,true)});cursor+=46+view.getUint16(cursor+28,true)+view.getUint16(cursor+30,true)+view.getUint16(cursor+32,true);}return {view,end,out};}

describe('办公 ZIP 包与 XML 安全边界',()=>{
 it('接受使用数据描述符的流式 ZIP',()=>{
  const chunks:Uint8Array[]=[];let done=false;
  const archive=new Zip((error,chunk,final)=>{if(error)throw error;chunks.push(chunk);done=final;});
  const file=new ZipDeflate('word/document.xml');archive.add(file);file.push(strToU8('<root>正文</root>'),true);archive.end();
  expect(done).toBe(true);const bytes=Buffer.concat(chunks);expect(readOfficePackage(bytes)['word/document.xml']).toEqual(strToU8('<root>正文</root>'));
 });
 it.each([0,6] as const)('接受存储/压缩方式 %s，保留原始部件字节',(level)=>{
  const bytes=zip({'word/document.xml':'<w:document><w:p>中文</w:p></w:document>','word/media/image.bin':'\u0000abc'},level);
  const files=readOfficePackage(bytes);expect(files['word/document.xml']).toEqual(strToU8('<w:document><w:p>中文</w:p></w:document>'));expect(files['word/media/image.bin']).toEqual(strToU8('\u0000abc'));
 });
 it.each(['../outside.xml','/absolute.xml','a/../b.xml','a\\b.xml','C:/file.xml','a//b.xml','a/./b.xml'])('拒绝路径 %s',(name)=>expect(()=>readOfficePackage(zip({[name]:'<root/>'}))).toThrow('路径'));
 it('拒绝重复大小写路径和截断包',()=>{
  expect(()=>readOfficePackage(zip({'word/document.xml':'<a/>','WORD/document.xml':'<b/>'}))).toThrow('重复路径');
  const bytes=zip({'word/document.xml':'<root/>'});expect(()=>readOfficePackage(bytes.subarray(0,bytes.length-1))).toThrow();
 });
 it('在解压之前检查条目数量、单件大小和总展开大小',()=>{
  const spy=vi.spyOn(Inflate.prototype,'push');
  try{
   const count=zip({'a.xml':'<a/>'});const c=entries(count);c.view.setUint16(c.end+8,2049,true);c.view.setUint16(c.end+10,2049,true);expect(()=>readOfficePackage(count)).toThrow('2048');
   const single=zip({'a.xml':'<a/>'});const s=entries(single);s.view.setUint32(s.out[0].central+24,OFFICE_PACKAGE_LIMITS.xmlBytes+1,true);expect(()=>readOfficePackage(single)).toThrow('8 MB');
   const total=zip(Object.fromEntries(Array.from({length:9},(_,i)=>[i+'.xml','<a/>'])));const t=entries(total);for(const entry of t.out){t.view.setUint32(entry.central+24,OFFICE_PACKAGE_LIMITS.xmlBytes,true);t.view.setUint32(entry.local+22,OFFICE_PACKAGE_LIMITS.xmlBytes,true);}expect(()=>readOfficePackage(total)).toThrow('64 MB');
   expect(spy).not.toHaveBeenCalled();
  }finally{spy.mockRestore();}
 });
 it('拒绝伪造展开大小、校验和错误与不一致头部',()=>{
  const forged=zip({'blob.bin':'a'.repeat(1000000)});const f=entries(forged);f.view.setUint32(f.out[0].central+24,1,true);f.view.setUint32(f.out[0].local+22,1,true);expect(()=>readOfficePackage(forged)).toThrow('实际展开大小');
  const corrupt=zip({'blob.bin':'abc'},0);const c=entries(corrupt);const position=c.out[0].local+30+c.view.getUint16(c.out[0].local+26,true)+c.view.getUint16(c.out[0].local+28,true);corrupt[position]^=1;expect(()=>readOfficePackage(corrupt)).toThrow('校验');
  const mismatch=zip({'blob.bin':'abc'});const m=entries(mismatch);m.view.setUint16(m.out[0].local+8,0,true);expect(()=>readOfficePackage(mismatch)).toThrow();
 });
 it('拒绝加密、ZIP64 与不支持的压缩方法',()=>{
  for(const [field,value]of [[8,1],[10,99],[24,0xffffffff]] as const){const bytes=zip({'blob.bin':'abc'}),record=entries(bytes);if(field===24)record.view.setUint32(record.out[0].central+field,value,true);else record.view.setUint16(record.out[0].central+field,value,true);expect(()=>readOfficePackage(bytes)).toThrow();}
 });
 it('严格解码 UTF 并检查所有 XML 部件中的 DTD',()=>{
  expect(readOfficeXml(encodeOfficeText('<root>中文</root>','utf-16le',true))).toContain('中文');
  expect(readOfficeXml(encodeOfficeText('<root>中文</root>','utf-16be',true))).toContain('中文');
  expect(()=>readOfficeXml(new Uint8Array([0xff,0xff,0xff]))).toThrow();
  expect(()=>readOfficeXml(strToU8('<?xml version="1.0" encoding="UTF-16"?><root/>'))).toThrow('编码');
  expect(()=>readOfficePackage(zip({'word/document.xml':'<root/>','unused.xml':'<!DOCTYPE root [<!ENTITY x SYSTEM "file:///etc/passwd">]><root>&x;</root>'}))).toThrow('DTD');
 });
 it('拒绝过深、错配和未闭合 XML；接受注释、CDATA 和含大于号的属性',()=>{
  expect(()=>readOfficeXml(strToU8('<x>'.repeat(129)+'</x>'.repeat(129)))).toThrow('深度');
  for(const xml of ['<a><b></a>','<a>','<a/><b/>','<a><!-- nope','<a>text</a>outside'])expect(()=>readOfficeXml(strToU8(xml))).toThrow();
  expect(readOfficeXml(strToU8('<?xml version="1.0"?><root note="a>b"><!-- hi --><![CDATA[<x>]]><child/></root>'))).toContain('CDATA');
 });
});
