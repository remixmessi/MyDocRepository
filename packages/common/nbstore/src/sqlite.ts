import {sqliteFetch,setStorageGeneration,setStorageAccount} from './desktop';
import {SCHEMA_VERSION} from './schema';
import type {DocumentStorage} from './index';
import type {AssetRecord,AssetStorage} from './assets';
export async function storageRequest(path:string,init?:RequestInit){const response=await sqliteFetch(path,init);if(!response.ok){let message='项目 SQLite 存储服务不可用';try{message=(await response.json()).error??message;}catch{/* Keep service error. */}throw new Error(message);}return response;}
function bytesInit(bytes:Uint8Array,method='POST'):RequestInit{return {method,headers:{'Content-Type':'application/octet-stream'},body:bytes.slice().buffer,keepalive:bytes.length<60000};}
export class SQLiteStorage implements DocumentStorage {
 async load(id:string){const response=await sqliteFetch('/document?id='+encodeURIComponent(id));if(response.status===204||response.status===404)return null;if(!response.ok)throw new Error('SQLite 文档读取失败');return new Uint8Array(await response.arrayBuffer());}
 async append(id:string,update:Uint8Array){await storageRequest('/document?id='+encodeURIComponent(id),bytesInit(update));}
 async remove(id:string){await storageRequest('/document?id='+encodeURIComponent(id),{method:'DELETE'});}
 close(){}
}
export class SQLiteAssetStorage implements AssetStorage {
 private path(workspace:string,id?:string){return '/asset?workspace='+encodeURIComponent(workspace)+(id?'&id='+encodeURIComponent(id):'');}
 async get(workspace:string,id:string):Promise<AssetRecord|null>{const response=await sqliteFetch(this.path(workspace,id));if(response.status===404)return null;if(!response.ok)throw new Error('SQLite 附件读取失败');const a=await response.json();return {id,name:a.name,mime:a.mime,size:a.size,bytes:Uint8Array.from(atob(a.data),c=>c.charCodeAt(0))};}
 async put(workspace:string,asset:AssetRecord){await storageRequest(this.path(workspace,asset.id)+'&name='+encodeURIComponent(asset.name)+'&mime='+encodeURIComponent(asset.mime),bytesInit(asset.bytes,'PUT'));}
 async remove(workspace:string,id:string){await storageRequest(this.path(workspace,id),{method:'DELETE'});}
 async removeWorkspace(workspace:string){await storageRequest('/assets?workspace='+encodeURIComponent(workspace),{method:'DELETE'});}
 close(){}
}
const pendingPreferences=new Set<Promise<void>>();
let preferenceChannel:BroadcastChannel|undefined;
export const projectPreferences={values:{} as Record<string,string>,getItem(key:string){return this.values[key]??null;},setItem(key:string,value:string){const operation=this.write(key,value);pendingPreferences.add(operation);void operation.then(()=>pendingPreferences.delete(operation),()=>pendingPreferences.delete(operation));return operation;},async flush(){await Promise.all([...pendingPreferences]);},async write(key:string,value:string){const previous=this.values[key];this.values[key]=value;try{const reply=await (await storageRequest('/preference?key='+encodeURIComponent(key),{method:'PUT',body:value,keepalive:value.length<60000})).json();if(this.values[key]===value)this.values[key]=reply.value??value;preferenceChannel?.postMessage({key,value:this.values[key]});window.dispatchEvent(new CustomEvent('mydoc-preference',{detail:{key,value:this.values[key]}}));}catch(error){if(this.values[key]===value){if(previous===undefined)delete this.values[key];else this.values[key]=previous;}throw error;}}};
export async function initializeProjectStorage(){const info=await (await storageRequest('/status')).json();if(!info.initialized||info.schemaVersion!==SCHEMA_VERSION)throw new Error('项目数据库尚未正确初始化，请重启本地服务');setStorageGeneration(info.generation);setStorageAccount(info.accountId);projectPreferences.values=info.preferences;if(typeof BroadcastChannel!=='undefined'&&!preferenceChannel){preferenceChannel=new BroadcastChannel('mydoc-project-preferences');preferenceChannel.onmessage=event=>{const {key,value}=event.data;projectPreferences.values[key]=value;window.dispatchEvent(new CustomEvent('mydoc-preference',{detail:{key,value}}));};}}
