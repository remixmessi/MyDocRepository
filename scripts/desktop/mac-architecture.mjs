import {open,readdir} from 'node:fs/promises';
import {join} from 'node:path';
const cpuNames=new Map([[0x01000007,'x64'],[0x0100000c,'arm64']]);
/** Inspect the Mach-O header, rather than trusting an artifact filename. */
export function machoArchitectures(bytes){
 const data=Buffer.from(bytes);if(data.length<8)return [];
 const magic=data.readUInt32BE(0);
 if(magic===0xcffaedfe||magic===0xcefaedfe)return [cpuNames.get(data.readUInt32LE(4))??'unknown'];
 if(magic===0xfeedfacf||magic===0xfeedface)return [cpuNames.get(data.readUInt32BE(4))??'unknown'];
 if(![0xcafebabe,0xbebafeca,0xcafebabf,0xbfbafeca].includes(magic))return [];
 const little=magic===0xbebafeca||magic===0xbfbafeca,read=offset=>little?data.readUInt32LE(offset):data.readUInt32BE(offset),count=read(4),width=[0xcafebabf,0xbfbafeca].includes(magic)?32:20;
 if(!count||count>32||data.length<8+count*width)throw new Error('无效的 Mach-O 架构表');
 return Array.from({length:count},(_,index)=>cpuNames.get(read(8+index*width))??'unknown');
}
export function assertArchitecture(actual,expected,path){if(!expected.every(arch=>actual.includes(arch)))throw new Error(`macOS 架构不匹配：${path}，目标 ${expected.join('+')}，实际 ${actual.join('+')||'非 Mach-O'}。不要把 arm64 electronDist 用于 Intel x64 打包；请使用对应架构的 Electron 运行时。`);}
async function header(path){const file=await open(path,'r');try{const bytes=Buffer.alloc(4096);const {bytesRead}=await file.read(bytes,0,bytes.length,0);return bytes.subarray(0,bytesRead);}finally{await file.close();}}
export async function verifyMacBundle(appPath,expected){
 let count=0;
 async function scan(dir){for(const entry of await readdir(dir,{withFileTypes:true})){const path=join(dir,entry.name);if(entry.isDirectory())await scan(path);else if(entry.isFile()){const actual=machoArchitectures(await header(path));if(actual.length||entry.name.endsWith('.node')){assertArchitecture(actual,expected,path);count++;}}}}
 await scan(join(appPath,'Contents','MacOS'));await scan(join(appPath,'Contents','Frameworks'));
 if(!count)throw new Error('应用中没有可验证的 macOS 可执行文件');
 return count;
}
