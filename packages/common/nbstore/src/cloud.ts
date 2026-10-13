import {DurableCloudQueue,type QueueSnapshot} from './sync-queue';
import {getStorageAccount} from './desktop';
import {Y,createDocumentAwareness,applyPresence,documentAwareness as importAwareness} from '@mydoc/crdt';import type {DocumentStorage} from './index';import type {AssetRecord,AssetStorage} from './assets';
export const CLOUD_ORIGIN=Symbol('mydoc-cloud');
let configuredCloudServer='';
export function cloudServerUrl(){return configuredCloudServer;}
export function configureCloudServer(value:string){
 const trimmed=value.trim().replace(/\/+$/,'');
 if(trimmed){const url=new URL(trimmed);if(!['http:','https:'].includes(url.protocol)||url.username||url.password)throw new Error('服务器地址必须是 http 或 https 地址');configuredCloudServer=url.origin+url.pathname.replace(/\/$/,'');}
 else configuredCloudServer='';
 return configuredCloudServer;
}
export interface CloudAccount {id:string;email:string;name:string;serverAdmin?:boolean;verified?:boolean;createdAt?:number;hasAvatar?:boolean;expiresAt?:number;session?:string;}
/** 邮件服务状态：未配置时界面要如实说明口令改为当场显示或由运维转交。 */
export interface CloudMailStatus {configured:boolean;reason?:string;}
export interface CloudWorkspace {id:string;name:string;createdAt:number;role:'owner'|'admin'|'member'|'viewer';}
export class CloudError extends Error {constructor(message:string,readonly status:number){super(message);}}
export function cloudUrl(path:string,base=configuredCloudServer){return base&&globalThis.window?.mydocDesktop?'/api/cloud-proxy/'+encodeURIComponent(base)+'/api/cloud'+path:base+'/api/cloud'+path;}
export async function cloudRequest(path:string,init?:RequestInit,server=configuredCloudServer){const response=await fetch(cloudUrl(path,server),{...init,credentials:'include',headers:{...Object.fromEntries(new Headers(init?.headers)),...(!server&&getStorageAccount()?{'X-Mydoc-Account':getStorageAccount()!}:{})}});if(!response.ok){let message='工作空间服务不可用';try{message=(await response.json()).error??message;}catch{}throw new CloudError(message,response.status);}return response;}
export async function cloudJson<T>(path:string,method='GET',value?:unknown,server=configuredCloudServer):Promise<T>{return (await cloudRequest(path,{method,...(value===undefined?{}:{headers:{'Content-Type':'application/json'},body:JSON.stringify(value)})},server)).json();}
export type SyncStatus='connecting'|'connected'|'offline'|'blocked';
const decode=(value:string)=>Uint8Array.from(atob(value),c=>c.charCodeAt(0));
export class CloudStorage implements DocumentStorage {
 private events?:EventSource;
 private docs=new Map<string,{doc:Y.Doc;error:(error:unknown)=>void;queue:Promise<void>}>();
 private listeners=new Set<()=>void>();private disposed=false;private synchronizing=0;private failed=new Set<string>();
 private readonly account=getStorageAccount();
 private readonly durable:DurableCloudQueue;
 private snapshots=new Map<string,QueueSnapshot>();
 private documentReadOnly=new Map<string,boolean>();
 isReadOnly(id:string){return this.readOnly||this.documentReadOnly.get(id)===true;}
 private sending=new Map<string,Promise<void>>();
 private retryTimer?:ReturnType<typeof setTimeout>;
 private undurable=new Set<string>();
 get memoryOnly(){return this.undurable.size>0;}
 get pendingCount(){return [...this.snapshots.values()].reduce((n,value)=>n+value.rows.length,0);}
 get lastSuccess(){return Math.max(0,...[...this.snapshots.values()].map(value=>value.lastSuccess));}
 private checkAccount(){if(getStorageAccount()!==this.account){const error=new CloudError('账户已变更，待提交队列已暂停',409);this.fail(error);throw error;}}
 status:SyncStatus='connecting';reason='';readOnly:boolean;needsLogin=false;
 private role:CloudWorkspace['role'];
 get canManageDocuments(){return !this.readOnly&&['owner','admin'].includes(this.role);}
 private online=()=>{if(this.status!=='blocked')void this.reconnect();};
 constructor(readonly workspace:string,role:CloudWorkspace['role']='member'){this.role=role;this.readOnly=role==='viewer';this.durable=new DurableCloudQueue(workspace,configuredCloudServer||globalThis.location?.origin||'http://localhost',this.account);globalThis.addEventListener?.('online',this.online);}
 async visibleDocuments(ids:string[]):Promise<Set<string>>{const unique=[...new Set(ids.filter(Boolean).map(id=>id.startsWith('doc:')?id:'doc:'+id))];if(!unique.length)return new Set();const query=unique.map(id=>'id='+encodeURIComponent(id)).join('&');const rows=await (await this.request(this.path('/document-visibility')+'&'+query)).json() as {id:string;allowed:boolean}[];return new Set(rows.filter(row=>row.allowed).map(row=>row.id));}
 subscribe=(listener:()=>void)=>{this.listeners.add(listener);return()=>{this.listeners.delete(listener);};};
 private notify(){for(const listener of this.listeners)listener();}
 private path(route:string,id?:string){return route+'?workspace='+encodeURIComponent(this.workspace)+(id?'&id='+encodeURIComponent(id):'');}
 private blocked(){return this.status==='blocked';}
 // 401 单独标记：会话到期只需重新登录，不是失去权限，界面据此给出重新登录入口而非刷新页面。
 private fail(error:unknown){if(this.disposed||this.blocked())return;if(error instanceof CloudError&&[401,403,409,410].includes(error.status)){this.status='blocked';this.readOnly=true;this.reason=error.message;this.needsLogin=error.status===401;this.events?.close();this.events=undefined;if(this.retryTimer)clearTimeout(this.retryTimer);}else{this.status='offline';this.reason=error instanceof Error?error.message:'连接中断';}this.notify();}
 private async request(path:string,init?:RequestInit,id?:string){try{this.checkAccount();return await cloudRequest(path,init);}catch(error){if(id?.startsWith('doc:')&&error instanceof CloudError&&error.status===403){this.documentReadOnly.set(id,true);this.status='offline';this.reason=error.message;this.notify();}else this.fail(error);throw error;}}
 private async refreshAccess(id:string){const access=await (await this.request(this.path('/access',id))).json() as {role:CloudWorkspace['role']};if(id.startsWith('doc:'))this.documentReadOnly.set(id,access.role==='viewer');else{this.role=access.role;this.readOnly=access.role==='viewer';}this.notify();}
 private async snapshot(id:string){this.checkAccount();const snapshot=await this.durable.list(id);this.snapshots.set(id,snapshot);this.notify();return snapshot;}
 async load(id:string){await this.refreshAccess(id);const response=await this.request(this.path('/document',id));const remote=response.status===204?null:new Uint8Array(await response.arrayBuffer());const snapshot=await this.snapshot(id);const updates=snapshot.rows.filter(row=>row.generation===snapshot.generation).map(row=>decode(row.data));return updates.length?Y.mergeUpdates([...(remote?[remote]:[]),...updates]):remote;}
 async append(id:string,update:Uint8Array){this.checkAccount();if(this.isReadOnly(id))throw new CloudError(this.reason||'该文档为只读，未提交内容可导出为草稿',403);const doc=this.docs.get(id)?.doc,snapshot=doc?Y.encodeStateAsUpdate(doc):update;try{await this.durable.append(id,this.undurable.has(id)?snapshot:update,snapshot);this.undurable.delete(id);}catch(error){this.undurable.add(id);throw error;}await this.snapshot(id);await this.drain(id).catch(error=>{this.failed.add(id);this.docs.get(id)?.error(error);});}
 private scheduleRetry(delay:number){if(this.retryTimer||this.disposed||this.blocked()||this.readOnly)return;this.retryTimer=setTimeout(()=>{this.retryTimer=undefined;for(const id of this.snapshots.keys())void this.drain(id).catch(error=>this.docs.get(id)?.error(error));},Math.max(1000,delay));}
 private drain(id:string,force=false):Promise<void>{const previous=this.sending.get(id)??Promise.resolve();const operation=previous.catch(()=>{}).then(async()=>{
  this.checkAccount();if(this.disposed||this.blocked()||this.isReadOnly(id))return;
  const snapshot=await this.snapshot(id);
  if(snapshot.rows.some(row=>row.generation!==snapshot.generation)){const error=new CloudError('数据库已恢复，旧代际草稿已保留，自动提交已暂停',409);this.fail(error);throw error;}
  for(const row of snapshot.rows){
   this.checkAccount();if(this.disposed||this.blocked()||this.isReadOnly(id))return;
   if(!force&&row.retryAt>Date.now()){this.status='offline';this.notify();this.scheduleRetry(row.retryAt-Date.now());return;}
   try{await this.request(this.path('/document',id),{method:'POST',headers:{'Content-Type':'application/octet-stream'},body:decode(row.data).slice().buffer},id);await this.durable.acknowledge(id,row.seq);}
   catch(error){await this.durable.fail(id,row.seq,error instanceof Error?error.message:'提交失败').catch(()=>{});await this.snapshot(id).catch(()=>{});if(!this.blocked())this.scheduleRetry(Math.min(60000,1000*2**Math.min(row.attempts,6)));throw error;}
  }
  const current=await this.snapshot(id);if(!current.rows.length){this.failed.delete(id);this.docs.get(id)?.error(undefined);if(!this.failed.size&&!this.synchronizing&&!this.blocked()){this.status='connected';this.reason='';}this.notify();}
 });this.sending.set(id,operation);void operation.finally(()=>{if(this.sending.get(id)===operation)this.sending.delete(id);}).catch(()=>{});return operation;}
 async remove(id:string){if(this.isReadOnly(id))throw new CloudError('该文档为只读',403);await this.request(this.path('/document',id),{method:'DELETE'},id);}
 private synchronize(id:string){const target=this.docs.get(id);if(!target)return Promise.resolve();
 target.queue=target.queue.catch(()=>{}).then(async()=>{if(this.disposed||this.status==='blocked')return;this.synchronizing++;this.status='connecting';this.notify();
 try{const response=await this.request(this.path('/sync',id),{method:'POST',body:Y.encodeStateVector(target.doc).slice().buffer});const value=await response.json() as {role:CloudWorkspace['role'];update:string;vector:string};
 if(this.disposed||this.blocked()||!this.docs.has(id))return;
 if(id.startsWith('doc:'))this.documentReadOnly.set(id,value.role==='viewer');else{this.role=value.role;this.readOnly=value.role==='viewer';}Y.applyUpdate(target.doc,decode(value.update),CLOUD_ORIGIN);
 if(!this.isReadOnly(id)){const update=Y.encodeStateAsUpdate(target.doc,decode(value.vector));if(update.length>2)await this.append(id,update);await this.drain(id);}
 if(this.blocked()||this.snapshots.get(id)?.rows.length)return;this.failed.delete(id);target.error(undefined);
 }catch(error){this.failed.add(id);target.error(error);}finally{this.synchronizing--;if(!this.disposed&&!this.blocked()&&!this.synchronizing){if(this.failed.size||this.pendingCount){this.status='offline';}else{this.status='connected';this.reason='';}}this.notify();}});return target.queue;
 }
 async reconnect(){if(this.disposed)return;try{const access=await (await this.request(this.path('/access'))).json() as {role:CloudWorkspace['role']};this.role=access.role;this.readOnly=access.role==='viewer';this.status='connecting';this.reason='';this.needsLogin=false;this.notify();if(!this.readOnly)for(const id of this.snapshots.keys()){await this.refreshAccess(id);await this.drain(id,true);}this.events?.close();this.events=undefined;this.startEvents();}catch(error){this.fail(error);}}
 private startEvents(){if(this.events||this.disposed||this.status==='blocked')return;
 const events=this.events=new EventSource(cloudUrl(this.path('/events','*')+(!configuredCloudServer?'&account='+encodeURIComponent(this.account??''):'')),{withCredentials:true});
 events.onmessage=event=>{try{const value=JSON.parse(event.data) as {id:string;data?:string;presence?:{client:number;clock:number;state:unknown}};const target=this.docs.get(value.id);if(target&&value.presence){const awareness=importAwareness(target.doc);if(awareness)applyPresence(awareness,value.presence.client,value.presence.clock,value.presence.state);}else if(target&&value.data)Y.applyUpdate(target.doc,decode(value.data),CLOUD_ORIGIN);}catch(error){for(const target of this.docs.values())target.error(error);}};
 events.onopen=()=>{this.status='connecting';this.notify();for(const id of this.docs.keys())void this.synchronize(id);};
 events.onerror=()=>{if(this.disposed||this.events!==events)return;this.status='offline';this.notify();void this.request(this.path('/access')).then(r=>r.json()).then((access:{role:CloudWorkspace['role']})=>{if(this.disposed||this.blocked())return;this.role=access.role;this.readOnly=access.role==='viewer';this.notify();}).catch(error=>this.fail(error));};
 }
 connect(id:string,doc:Y.Doc,onError:(error:unknown)=>void){const awareness=createDocumentAwareness(doc);this.docs.set(id,{doc,error:onError,queue:Promise.resolve()});
 const send=()=>{if(this.status==='blocked')return;const state=awareness.getLocalState();void this.request(this.path('/presence',id),{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({client:doc.clientID,clock:awareness.meta.get(doc.clientID)?.clock??0,cursor:state?.cursor??null,active:!!state})}).catch(()=>{});};
 const changed=(_change:unknown,origin:unknown)=>{if(origin!=='remote')send();};awareness.on('update',changed);awareness.setLocalStateField('user',{name:'我',color:'#3159cb'});
 this.startEvents();if(this.events?.readyState===EventSource.OPEN)void this.synchronize(id);
 return()=>{awareness.setLocalState(null);awareness.off('update',changed);awareness.destroy();this.docs.delete(id);this.failed.delete(id);if(!this.docs.size){this.events?.close();this.events=undefined;}};
 }
 close(){globalThis.removeEventListener?.('online',this.online);this.disposed=true;if(this.retryTimer)clearTimeout(this.retryTimer);this.events?.close();this.docs.clear();this.listeners.clear();}
}
export class CloudAssets implements AssetStorage {
 constructor(readonly workspace:string){}private path(id?:string){return (id?'/asset':'/assets')+'?workspace='+encodeURIComponent(this.workspace)+(id?'&id='+encodeURIComponent(id):'');}
 async get(_workspace:string,id:string):Promise<AssetRecord|null>{const response=await fetch(cloudUrl(this.path(id)),{credentials:'include',headers:!configuredCloudServer&&getStorageAccount()?{'X-Mydoc-Account':getStorageAccount()!}:{}});if(response.status===404)return null;if(!response.ok)throw new Error('共享附件读取失败');const a=await response.json();return {...a,bytes:Uint8Array.from(atob(a.data),c=>c.charCodeAt(0))};}
 async getForDocument(_workspace:string,id:string,doc:string){return this.getScoped(id,doc);}
 private async getScoped(id:string,doc:string):Promise<AssetRecord|null>{const response=await fetch(cloudUrl(this.path(id)+'&doc='+encodeURIComponent(doc)),{credentials:'include',headers:!configuredCloudServer&&getStorageAccount()?{'X-Mydoc-Account':getStorageAccount()!}:{}});if(response.status===404)return null;if(!response.ok)throw new Error('共享附件读取失败');const a=await response.json();return {...a,bytes:Uint8Array.from(atob(a.data),c=>c.charCodeAt(0))};}
 async put(_workspace:string,a:AssetRecord){await cloudRequest(this.path(a.id)+'&name='+encodeURIComponent(a.name)+'&mime='+encodeURIComponent(a.mime),{method:'PUT',body:a.bytes.slice().buffer});}
 async putForDocument(_workspace:string,a:AssetRecord,doc:string){await cloudRequest(this.path(a.id)+'&name='+encodeURIComponent(a.name)+'&mime='+encodeURIComponent(a.mime)+'&doc='+encodeURIComponent(doc),{method:'PUT',body:a.bytes.slice().buffer});}
 async remove(_workspace:string,id:string){await cloudRequest(this.path(id),{method:'DELETE'});}
 async removeForDocument(_workspace:string,id:string,doc:string){await cloudRequest(this.path(id)+'&doc='+encodeURIComponent(doc),{method:'DELETE'});}
 async removeWorkspace(_workspace:string){await cloudRequest(this.path(),{method:'DELETE'});}
 close(){}
}
