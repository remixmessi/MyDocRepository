import {dirname,resolve,sep} from 'node:path';
import {copyFileSync,existsSync,mkdirSync,renameSync,unlinkSync} from 'node:fs';
import type {DesktopStorageRequest} from '../../../../common/nbstore/src/desktop';
export function safeExternalUrl(value:string){try{const u=new URL(value);return ['https:','http:'].includes(u.protocol)&&!u.username&&!u.password?u.href:null;}catch{return null;}}
export function validStorageRequest(value:unknown):value is DesktopStorageRequest {
 if(!value||typeof value!=='object')return false;
 const r=value as DesktopStorageRequest;
 if(r.generation!==undefined&&(typeof r.generation!=='string'||r.generation.length>100))return false;
 if(r.contentType!==undefined&&(typeof r.contentType!=='string'||r.contentType.length>200))return false;
 if(typeof r.path!=='string'||r.path.length>10000||!r.path.startsWith('/')||r.path.startsWith('//')||!['GET','POST','PUT','PATCH','DELETE'].includes(r.method))return false;
 const u=new URL(r.path,'http://storage');
 if(u.origin!=='http://storage'||u.hash)return false;
 const methods:Record<string,string[]>={'/local-account':['POST'],'/status':['GET'],'/backups':['GET'],'/backups/create':['POST'],'/backups/settings':['PUT'],'/backups/preview':['GET'],'/backups/restore':['POST'],'/sync-drafts':['GET'],'/sync-queue':['GET','POST','PATCH','DELETE'],'/document':['GET','POST','DELETE'],'/asset':['GET','PUT','DELETE'],'/assets':['DELETE'],'/preference':['PUT']};
 return !!methods[u.pathname]?.includes(r.method)&&(r.body===undefined||r.body instanceof Uint8Array&&r.body.byteLength<=64*1024*1024);
}
export function staticFile(root:string,url:string){try{const p=decodeURIComponent(new URL(url,'http://local').pathname);if(p.includes('\0')||p.includes('\\'))return null;const file=resolve(root,'.'+(p==='/'?'/index.html':p));return file.startsWith(resolve(root)+sep)?file:null;}catch{return null;}}
/** User data must live outside the installed application, which is replaced during upgrades. */
export function desktopDatabasePath(packaged:boolean,projectRoot:string,userDataDirectory:string){return resolve(packaged?userDataDirectory:projectRoot,'data/mydoc.sqlite');}

/** Copy databases created by older versions that stored them beside the executable. */
export function migrateLegacyDesktopDatabase(applicationDirectory:string,userDataDirectory:string){
 const legacy=resolve(applicationDirectory,'data/mydoc.sqlite');
 const target=resolve(userDataDirectory,'data/mydoc.sqlite');
 if(existsSync(target)||!existsSync(legacy))return false;
 mkdirSync(dirname(target),{recursive:true});
 const temp=target+'.migrating-'+process.pid;
 const sidecars=['-wal','-shm'];
 try{
  copyFileSync(legacy,temp);
  for(const suffix of sidecars)if(existsSync(legacy+suffix))copyFileSync(legacy+suffix,temp+suffix);
  renameSync(temp,target);
  for(const suffix of sidecars)if(existsSync(temp+suffix))renameSync(temp+suffix,target+suffix);
  return true;
 }catch(error){
  for(const path of [temp,...sidecars.map(suffix=>temp+suffix)])try{unlinkSync(path);}catch{}
  throw error;
 }
}
