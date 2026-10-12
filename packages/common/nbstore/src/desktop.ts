export interface DesktopStorageRequest {path:string;method:string;body?:Uint8Array;generation?:string;contentType?:string;}
export interface DesktopStorageReply {status:number;contentType:string;body:Uint8Array;}
export interface DesktopOpenFiles {id:string;files:{name:string;bytes:Uint8Array;modified:number}[];errors:string[];}
export interface DesktopBridge {
 readonly platform:string;
 updateStatus():Promise<import('./updates').DesktopUpdateState>;
 updateAction(action:'check'|'download'|'install'):Promise<import('./updates').DesktopUpdateState>;
 setAutomaticUpdateChecks(value:boolean):Promise<import('./updates').DesktopUpdateState>;
 onUpdate(callback:(state:import('./updates').DesktopUpdateState)=>void):()=>void;
 onShowUpdates(callback:()=>void):()=>void;
 chooseBackupDirectory():Promise<import('./backups').BackupStatus|undefined>;
 openBackupDirectory():Promise<void>;
 takeDocumentLink():Promise<{id:string;target?:import('./deep-link').DocumentLocation;error?:string}|undefined>;
 acknowledgeDocumentLink(id:string):Promise<void>;
 onDocumentLink(callback:()=>void):()=>void;
 storage(request:DesktopStorageRequest):Promise<DesktopStorageReply>;
 takeOpenFiles():Promise<DesktopOpenFiles|undefined>;
 acknowledgeOpenFiles(id:string):Promise<void>;
 onOpenFiles(callback:()=>void):()=>void;
 onPrepareClose(callback:()=>Promise<void>):()=>void;
}
declare global {interface Window {mydocDesktop?:DesktopBridge;}}
let storageGeneration:string|undefined;let storageAccount:string|undefined;
export function setStorageAccount(value?:string){storageAccount=value;}
export function getStorageAccount(){return storageAccount;}
export function setStorageGeneration(value:string){storageGeneration=value;}
export async function sqliteFetch(path:string,init?:RequestInit):Promise<Response>{
 if(!window.mydocDesktop)return fetch('/api/storage'+path,{...init,headers:{...Object.fromEntries(new Headers(init?.headers)),...(storageGeneration?{'X-Mydoc-Generation':storageGeneration}:{}),...(storageAccount?{'X-Mydoc-Account':storageAccount}:{})}});
 const body=init?.body===undefined?undefined:new Uint8Array(await new Response(init.body).arrayBuffer());
 const contentType=new Headers(init?.headers).get('Content-Type')??undefined;
 const reply=await window.mydocDesktop.storage({path,method:init?.method??'GET',body,generation:storageGeneration,contentType});
 return new Response([204,304].includes(reply.status)?null:reply.body.slice().buffer,{status:reply.status,headers:{'Content-Type':reply.contentType}});
}
