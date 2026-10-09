import {getStorageAccount} from './desktop';
import {Y,createDocumentAwareness,applyPresence,documentAwareness as importAwareness} from '@mydoc/crdt';import type {DocumentStorage} from './index';import type {AssetRecord,AssetStorage} from './assets';
export const CLOUD_ORIGIN=Symbol('mydoc-cloud');
export interface CloudAccount {id:string;email:string;name:string;}
export interface CloudWorkspace {id:string;name:string;createdAt:number;role:'owner'|'admin'|'member'|'viewer';}
export class CloudError extends Error {constructor(message:string,readonly status:number){super(message);}}
export async function cloudRequest(path:string,init?:RequestInit){const response=await fetch('/api/cloud'+path,{...init,headers:{...Object.fromEntries(new Headers(init?.headers)),...(getStorageAccount()?{'X-Mydoc-Account':getStorageAccount()!}:{})}});if(!response.ok){let message='工作空间服务不可用';try{message=(await response.json()).error??message;}catch{}throw new CloudError(message,response.status);}return response;}
export async function cloudJson<T>(path:string,method='GET',value?:unknown):Promise<T>{return (await cloudRequest(path,{method,...(value===undefined?{}:{headers:{'Content-Type':'application/json'},body:JSON.stringify(value)})})).json();}
export type SyncStatus='connecting'|'connected'|'offline'|'blocked';
const decode=(value:string)=>Uint8Array.from(atob(value),c=>c.charCodeAt(0));
export class CloudStorage implements DocumentStorage {
 private events?:EventSource;
 private docs=new Map<string,{doc:Y.Doc;error:(error:unknown)=>void;queue:Promise<void>}>();
 private listeners=new Set<()=>void>();private disposed=false;private synchronizing=0;private failed=new Set<string>();
 status:SyncStatus='connecting';reason='';readOnly:boolean;
 private online=()=>{if(this.status!=='blocked')void this.reconnect();};
 constructor(readonly workspace:string,role:CloudWorkspace['role']='member'){this.readOnly=role==='viewer';globalThis.addEventListener?.('online',this.online);}
 subscribe=(listener:()=>void)=>{this.listeners.add(listener);return()=>{this.listeners.delete(listener);};};
 private notify(){for(const listener of this.listeners)listener();}
 private path(route:string,id?:string){return route+'?workspace='+encodeURIComponent(this.workspace)+(id?'&id='+encodeURIComponent(id):'');}
 private blocked(){return this.status==='blocked';}
 private fail(error:unknown){if(this.disposed||this.blocked())return;if(error instanceof CloudError&&[401,403,409].includes(error.status)){this.status='blocked';this.readOnly=true;this.reason=error.message;this.events?.close();this.events=undefined;}else{this.status='offline';this.reason='连接中断，改动待提交；请保留页面。';}this.notify();}
 private async request(path:string,init?:RequestInit){try{return await cloudRequest(path,init);}catch(error){this.fail(error);throw error;}}
 async load(id:string){const response=await this.request(this.path('/document',id));return response.status===204?null:new Uint8Array(await response.arrayBuffer());}
 async append(id:string,update:Uint8Array){if(this.readOnly)throw new CloudError(this.reason||'查看者只能阅读工作空间',403);try{await this.request(this.path('/document',id),{method:'POST',headers:{'Content-Type':'application/octet-stream'},body:update.slice().buffer});}catch(error){this.failed.add(id);throw error;}}
 async remove(id:string){await this.request(this.path('/document',id),{method:'DELETE'});}
 private synchronize(id:string){const target=this.docs.get(id);if(!target)return Promise.resolve();
 target.queue=target.queue.catch(()=>{}).then(async()=>{if(this.disposed||this.status==='blocked')return;this.synchronizing++;this.status='connecting';this.notify();
 try{const response=await this.request(this.path('/sync',id),{method:'POST',body:Y.encodeStateVector(target.doc).slice().buffer});const value=await response.json() as {role:CloudWorkspace['role'];update:string;vector:string};
 if(this.disposed||this.blocked()||!this.docs.has(id))return;
 this.readOnly=value.role==='viewer';Y.applyUpdate(target.doc,decode(value.update),CLOUD_ORIGIN);
 if(!this.readOnly){const update=Y.encodeStateAsUpdate(target.doc,decode(value.vector));await this.append(id,update);}
 if(this.blocked())return;this.failed.delete(id);target.error(undefined);
 }catch(error){this.failed.add(id);target.error(error);}finally{this.synchronizing--;if(!this.disposed&&!this.blocked()&&!this.synchronizing){if(this.failed.size){this.status='offline';}else{this.status='connected';this.reason='';}}this.notify();}});return target.queue;
 }
 async reconnect(){if(this.disposed)return;try{const access=await cloudJson<{role:CloudWorkspace['role']}>(this.path('/access'));this.readOnly=access.role==='viewer';this.status='connecting';this.reason='';this.notify();this.events?.close();this.events=undefined;this.startEvents();}catch(error){this.fail(error);}}
 private startEvents(){if(this.events||this.disposed||this.status==='blocked')return;
 const events=this.events=new EventSource('/api/cloud'+this.path('/events','*')+'&account='+encodeURIComponent(getStorageAccount()??''));
 events.onmessage=event=>{try{const value=JSON.parse(event.data) as {id:string;data?:string;presence?:{client:number;clock:number;state:unknown}};const target=this.docs.get(value.id);if(target&&value.presence){const awareness=importAwareness(target.doc);if(awareness)applyPresence(awareness,value.presence.client,value.presence.clock,value.presence.state);}else if(target&&value.data)Y.applyUpdate(target.doc,decode(value.data),CLOUD_ORIGIN);}catch(error){for(const target of this.docs.values())target.error(error);}};
 events.onopen=()=>{this.status='connecting';this.notify();for(const id of this.docs.keys())void this.synchronize(id);};
 events.onerror=()=>{if(this.disposed||this.events!==events)return;this.status='offline';this.notify();void cloudJson<{role:CloudWorkspace['role']}>(this.path('/access')).then(access=>{if(this.disposed||this.blocked())return;this.readOnly=access.role==='viewer';this.notify();}).catch(error=>this.fail(error));};
 }
 connect(id:string,doc:Y.Doc,onError:(error:unknown)=>void){const awareness=createDocumentAwareness(doc);this.docs.set(id,{doc,error:onError,queue:Promise.resolve()});
 const send=()=>{if(this.status==='blocked')return;const state=awareness.getLocalState();void cloudJson(this.path('/presence',id),'POST',{client:doc.clientID,clock:awareness.meta.get(doc.clientID)?.clock??0,cursor:state?.cursor??null,active:!!state}).catch(()=>{});};
 const changed=(_change:unknown,origin:unknown)=>{if(origin!=='remote')send();};awareness.on('update',changed);awareness.setLocalStateField('user',{name:'我',color:'#3159cb'});
 this.startEvents();if(this.events?.readyState===EventSource.OPEN)void this.synchronize(id);
 return()=>{awareness.setLocalState(null);awareness.off('update',changed);awareness.destroy();this.docs.delete(id);this.failed.delete(id);if(!this.docs.size){this.events?.close();this.events=undefined;}};
 }
 close(){globalThis.removeEventListener?.('online',this.online);this.disposed=true;this.events?.close();this.docs.clear();this.listeners.clear();}
}
export class CloudAssets implements AssetStorage {
 constructor(readonly workspace:string){}private path(id?:string){return (id?'/asset':'/assets')+'?workspace='+encodeURIComponent(this.workspace)+(id?'&id='+encodeURIComponent(id):'');}
 async get(_workspace:string,id:string):Promise<AssetRecord|null>{const response=await fetch('/api/cloud'+this.path(id),{headers:getStorageAccount()?{'X-Mydoc-Account':getStorageAccount()!}:{}});if(response.status===404)return null;if(!response.ok)throw new Error('共享附件读取失败');const a=await response.json();return {...a,bytes:Uint8Array.from(atob(a.data),c=>c.charCodeAt(0))};}
 async put(_workspace:string,a:AssetRecord){await cloudRequest(this.path(a.id)+'&name='+encodeURIComponent(a.name)+'&mime='+encodeURIComponent(a.mime),{method:'PUT',body:a.bytes.slice().buffer});}
 async remove(_workspace:string,id:string){await cloudRequest(this.path(id),{method:'DELETE'});}
 async removeWorkspace(_workspace:string){await cloudRequest(this.path(),{method:'DELETE'});}
 close(){}
}
