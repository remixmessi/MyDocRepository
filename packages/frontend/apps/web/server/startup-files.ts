import Database from 'better-sqlite3';
import {createHash,randomUUID} from 'node:crypto';
import {closeSync,copyFileSync,existsSync,fsyncSync,mkdirSync,mkdtempSync,openSync,readSync,renameSync,rmSync,statSync,writeFileSync} from 'node:fs';
import {dirname,join} from 'node:path';

const suffixes=['','-wal','-shm','-journal'];
function syncFile(path:string){const fd=openSync(path,'r');try{fsyncSync(fd);}finally{closeSync(fd);}}
function checksum(path:string){
 const hash=createHash('sha256'),fd=openSync(path,'r'),buffer=Buffer.allocUnsafe(1024*1024);
 try{for(;;){const bytes=readSync(fd,buffer,0,buffer.length,null);if(!bytes)break;hash.update(buffer.subarray(0,bytes));}}finally{closeSync(fd);}
 return hash.digest('hex');
}
function identities(path:string){
 const db=new Database(path,{readonly:true,fileMustExist:true});
 try{
  if(!db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name='cloud_users'").get())return [];
  return (db.prepare('SELECT id FROM cloud_users').all() as {id:string}[]).map(row=>row.id);
 }finally{db.close();}
}

/** Run before any SQLite connection opens the original database or recovers WAL. */
export function protectIncomingWal(path:string){
 if(!existsSync(path+'-wal')||statSync(path+'-wal').size===0)return;
 const original=new Map(suffixes.flatMap(suffix=>existsSync(path+suffix)?[[suffix,statSync(path+suffix,{bigint:true})] as const]:[]));
 const directory=join(dirname(path),'upgrade-inputs');mkdirSync(directory,{recursive:true,mode:0o700});
 const stage=join(directory,'.partial-'+randomUUID());mkdirSync(stage,{mode:0o700});
 let probe:string|undefined;
 try{
  const files=suffixes.filter(suffix=>original.has(suffix)).map(suffix=>{
   const name='database.sqlite'+suffix,target=join(stage,name);copyFileSync(path+suffix,target);syncFile(target);
   return {name,bytes:Number(original.get(suffix)!.size),sha256:checksum(target)};
  });
  for(const suffix of suffixes){
   const before=original.get(suffix),after=existsSync(path+suffix)?statSync(path+suffix,{bigint:true}):undefined;
   if(!!before!==!!after||before&&after&&(before.size!==after.size||before.mtimeNs!==after.mtimeNs||before.ctimeNs!==after.ctimeNs))throw new Error('数据库文件在启动检查期间发生变化，请先停止其他使用该数据目录的进程');
  }
  const key=createHash('sha256').update(JSON.stringify(files)).digest('hex').slice(0,24);
  let saved=join(directory,'input-'+key);
  writeFileSync(join(stage,'manifest.json'),JSON.stringify({createdAt:new Date().toISOString(),source:path,files},null,2),{mode:0o600,flag:'wx'});syncFile(join(stage,'manifest.json'));
  if(existsSync(saved)&&!files.every(file=>existsSync(join(saved,file.name))&&checksum(join(saved,file.name))===file.sha256))saved+='-'+randomUUID();
  if(existsSync(saved))rmSync(stage,{recursive:true});else renameSync(stage,saved);
  console.log('MyDoc original SQLite files preserved before WAL recovery: '+saved);
  probe=mkdtempSync(join(directory,'.probe-'));
  const main=join(probe,'main.sqlite'),combined=join(probe,'combined.sqlite');
  copyFileSync(join(saved,'database.sqlite'),main);
  copyFileSync(join(saved,'database.sqlite'),combined);
  copyFileSync(join(saved,'database.sqlite-wal'),combined+'-wal');
  // A valid interrupted checkpoint may leave the standalone main file unreadable.
  // In that case keep the raw evidence and let SQLite validate the complete pair.
  let base:string[]|undefined;try{base=identities(main);}catch{base=undefined;}
  const recovered=new Set(identities(combined));
  // The application never deletes/reassigns account IDs. Losing an ID merely by
  // applying WAL indicates an unsafe restore or external destructive write.
  if(base?.some(id=>!recovered.has(id)))throw new Error(`WAL 应用后已有账号消失，疑似恢复主库时混入了另一份 WAL。已停止启动，未打开或修改原数据库；原始文件已保留在 ${saved}。请将正确的一组数据库文件恢复到新的空目录。`);
 }finally{
  if(probe)rmSync(probe,{recursive:true,force:true});
  rmSync(stage,{recursive:true,force:true});
 }
}
