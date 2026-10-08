import {resolve,sep} from 'node:path';
import type {DesktopStorageRequest} from '../../../../common/nbstore/src/desktop';
export function safeExternalUrl(value:string){try{const u=new URL(value);return ['https:','http:'].includes(u.protocol)&&!u.username&&!u.password?u.href:null;}catch{return null;}}
export function validStorageRequest(value:unknown):value is DesktopStorageRequest {
 if(!value||typeof value!=='object')return false;
 const r=value as DesktopStorageRequest;
 if(typeof r.path!=='string'||r.path.length>10000||!r.path.startsWith('/')||r.path.startsWith('//')||!['GET','POST','PUT','DELETE'].includes(r.method))return false;
 const u=new URL(r.path,'http://storage');
 if(u.origin!=='http://storage'||u.hash)return false;
 const methods:Record<string,string[]>={'/status':['GET'],'/document':['GET','POST','DELETE'],'/asset':['GET','PUT','DELETE'],'/assets':['DELETE'],'/preference':['PUT']};
 return !!methods[u.pathname]?.includes(r.method)&&(r.body===undefined||r.body instanceof Uint8Array&&r.body.byteLength<=64*1024*1024);
}
export function staticFile(root:string,url:string){try{const p=decodeURIComponent(new URL(url,'http://local').pathname);if(p.includes('\0')||p.includes('\\'))return null;const file=resolve(root,'.'+(p==='/'?'/index.html':p));return file.startsWith(resolve(root)+sep)?file:null;}catch{return null;}}
export function desktopDatabasePath(packaged:boolean,projectRoot:string,userData:string){return resolve(packaged?userData:projectRoot,'data/mydoc.sqlite');}
