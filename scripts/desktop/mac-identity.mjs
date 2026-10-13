import {createHash} from 'node:crypto';
import {readFile,readdir,writeFile} from 'node:fs/promises';
import {join,relative} from 'node:path';

// Electron's stock executable UUID is shared by otherwise unrelated apps.
// Namespace it before signing so macOS network privacy can distinguish MyDoc.
export function namespaceExecutableUUID(bytes,identity){
 const data=Buffer.from(bytes),locations=[];
 function image(start,size){
  const end=start+size;
  if(size<32||end>data.length)throw new Error('Invalid Mach-O image');
  const magic=data.readUInt32BE(start),little=magic===0xcffaedfe;
  if(!little&&magic!==0xfeedfacf)throw new Error('Expected a 64-bit Mach-O executable');
  const read=offset=>little?data.readUInt32LE(offset):data.readUInt32BE(offset);
  const count=read(start+16),limit=start+32+read(start+20);
  if(limit>end||count>(limit-start-32)/8)throw new Error('Invalid Mach-O load commands');
  let offset=start+32,uuids=0;
  for(let i=0;i<count;i++){
   if(offset+8>limit)throw new Error('Truncated Mach-O load command');
   const command=read(offset),width=read(offset+4);
   if(width<8||offset+width>limit)throw new Error('Invalid Mach-O load command size');
   if(command===0x1b){if(width!==24)throw new Error('Invalid LC_UUID');locations.push(offset+8);uuids++;}
   offset+=width;
  }
  if(offset!==limit||uuids!==1)throw new Error('Expected exactly one LC_UUID per architecture');
 }
 if(data.length<8)throw new Error('Truncated Mach-O executable');
 const magic=data.readUInt32BE(0);
 if([0xcafebabe,0xbebafeca,0xcafebabf,0xbfbafeca].includes(magic)){
  const little=magic===0xbebafeca||magic===0xbfbafeca,wide=magic===0xcafebabf||magic===0xbfbafeca;
  const read=offset=>little?data.readUInt32LE(offset):data.readUInt32BE(offset);
  const read64=offset=>Number(little?data.readBigUInt64LE(offset):data.readBigUInt64BE(offset));
  const count=read(4),width=wide?32:20,tableEnd=8+count*width,ranges=[];
  if(!count||count>32||tableEnd>data.length)throw new Error('Invalid Mach-O architecture table');
  for(let i=0;i<count;i++){
   const entry=8+i*width,start=wide?read64(entry+8):read(entry+8),size=wide?read64(entry+16):read(entry+12);
   if(!Number.isSafeInteger(start+size)||start<tableEnd||ranges.some(([a,b])=>start<b&&start+size>a))throw new Error('Invalid Mach-O architecture range');
   image(start,size);ranges.push([start,start+size]);
  }
 }else image(0,data.length);
 for(const offset of locations){
  const uuid=createHash('sha256').update(identity).update('\0').update(data.subarray(offset,offset+16)).digest().subarray(0,16);
  uuid[6]=(uuid[6]&0x0f)|0x50;uuid[8]=(uuid[8]&0x3f)|0x80;
  uuid.copy(data,offset);
 }
 return data;
}

export async function prepareMacIdentity(appPath,appId,executable){
 const binaries=[join(appPath,'Contents','MacOS',executable)],frameworks=join(appPath,'Contents','Frameworks');
 for(const entry of await readdir(frameworks,{withFileTypes:true})){
  if(!entry.isDirectory()||!entry.name.endsWith('.app'))continue;
  const dir=join(frameworks,entry.name,'Contents','MacOS');
  for(const binary of await readdir(dir,{withFileTypes:true}))if(binary.isFile())binaries.push(join(dir,binary.name));
 }
 for(const path of binaries){const bytes=await readFile(path);await writeFile(path,namespaceExecutableUUID(bytes,appId+'/'+relative(appPath,path)));}
 return binaries.length;
}
