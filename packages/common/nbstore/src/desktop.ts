export interface DesktopStorageRequest {path:string;method:string;body?:Uint8Array;}
export interface DesktopStorageReply {status:number;contentType:string;body:Uint8Array;}
export interface DesktopBridge {
 readonly platform:string;
 storage(request:DesktopStorageRequest):Promise<DesktopStorageReply>;
 onPrepareClose(callback:()=>Promise<void>):()=>void;
}
declare global {interface Window {mydocDesktop?:DesktopBridge;}}
export async function sqliteFetch(path:string,init?:RequestInit):Promise<Response>{
 if(!window.mydocDesktop)return fetch('/api/storage'+path,init);
 const body=init?.body===undefined?undefined:new Uint8Array(await new Response(init.body).arrayBuffer());
 const reply=await window.mydocDesktop.storage({path,method:init?.method??'GET',body});
 return new Response([204,304].includes(reply.status)?null:reply.body.slice().buffer,{status:reply.status,headers:{'Content-Type':reply.contentType}});
}
