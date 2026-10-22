import {basename,extname,isAbsolute,resolve} from 'node:path';
import {open} from 'node:fs/promises';
import {randomUUID} from 'node:crypto';
import type {DesktopOpenFiles} from '../../../../common/nbstore/src/desktop';
import {OFFICE_IMPORT_EXTENSIONS,MAX_OFFICE_FILE_BYTES,officeImportable} from '../../../../common/nbstore/src/office';
export const OPEN_FILE_EXTENSIONS=['md','markdown','txt','zip',...OFFICE_IMPORT_EXTENSIONS];
const openable=(path:string)=>OPEN_FILE_EXTENSIONS.includes(extname(path).slice(1).toLowerCase());
export const OPEN_FILE_LIMITS={files:100,bytes:50*1024*1024,batches:8};
export function commandLineFiles(argv:string[],cwd:string,packaged:boolean){
 return argv.slice(packaged?1:2).filter(value=>!value.startsWith('-')&&!/^mydoc:/i.test(value)&&openable(value)).map(value=>isAbsolute(value)?value:resolve(cwd,value));
}
// Only OS open-file events, launch arguments and the native picker may enqueue paths.
// The renderer can consume queued bytes; it cannot supply an arbitrary read path.
export class OpenFileQueue {
 private batches:{id:string;paths:string[];error?:string;result?:Promise<DesktopOpenFiles>}[]=[];
 enqueue(paths:string[]){
  const queued=new Set(this.batches.flatMap(batch=>batch.paths));
  const unique=[...new Set(paths.map(path=>resolve(path)))].filter(path=>!queued.has(path));
  if(!unique.length)return false;
  // Office files open the office import; everything else goes to the document import.
  const groups=[unique.filter(path=>!officeImportable(basename(path))),unique.filter(path=>officeImportable(basename(path)))].filter(group=>group.length);
  if(this.batches.length+groups.length>OPEN_FILE_LIMITS.batches)return false;
  for(const group of groups)this.batches.push({id:randomUUID(),paths:group.slice(0,OPEN_FILE_LIMITS.files),error:group.length>OPEN_FILE_LIMITS.files?'每次最多打开 100 个文件，超出部分未读取。':undefined});return true;
 }
 async take():Promise<DesktopOpenFiles|undefined>{
  const batch=this.batches[0];if(!batch)return;
  batch.result??=(async()=>{
   const result:DesktopOpenFiles={id:batch.id,files:[],errors:batch.error?[batch.error]:[]};let total=0;
   for(const path of batch.paths){let file;const name=basename(path);try{
    if(!openable(name))throw new Error('仅支持 Markdown、文本、ZIP 和 Word / Excel / PPT / PDF；图片和附件请拖入文档或白板。');
    file=await open(path,'r');const stat=await file.stat();if(!stat.isFile())throw new Error('请打开文件；文件夹请在导入窗口选择。');
    if(officeImportable(name)&&stat.size>MAX_OFFICE_FILE_BYTES)throw new Error('办公文件原件超过 20 MB');
    if(stat.size>OPEN_FILE_LIMITS.bytes-total)throw new Error('所选文件合计不能超过 50 MB');
    // Bound the allocation and read, including a byte to detect growth after stat.
    const bytes=Buffer.alloc(stat.size+1);let length=0;while(length<bytes.length){const read=await file.read(bytes,length,bytes.length-length,length);if(!read.bytesRead)break;length+=read.bytesRead;}
    if(length>stat.size)throw new Error('文件正在变化，请保存文件后重新打开。');
    total+=length;result.files.push({name,bytes:new Uint8Array(bytes.subarray(0,length)),modified:stat.mtimeMs});
   }catch(error){const code=(error as NodeJS.ErrnoException).code;result.errors.push(`${name}：${code==='ENOENT'?'文件不存在，请重新选择':code==='EACCES'||code==='EPERM'?'没有读取权限':code?'读取失败，请重新选择':error instanceof Error?error.message:'读取失败'}`);}finally{await file?.close();}}
   return result;
  })();return batch.result;
 }
 acknowledge(id:unknown){if(typeof id==='string'&&this.batches[0]?.id===id)this.batches.shift();}
 get pending(){return this.batches.length;}
}
