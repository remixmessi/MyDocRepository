import {sqliteFetch,getStorageAccount} from './desktop';
export interface PendingUpdate {seq:number;data:string;attempts:number;retryAt:number;error:string|null;generation:string;}
export interface QueueSnapshot {generation:string;rows:PendingUpdate[];lastSuccess:number;}
export class DurableCloudQueue {
 constructor(readonly workspace:string,readonly server:string,readonly account:string|undefined){}
 async request(id:string,method='GET',body?:BodyInit):Promise<Response>{
  if(getStorageAccount()!==this.account)throw new Error('账户已变更，待提交队列已暂停');
  const path='/sync-queue?'+new URLSearchParams({server:this.server,workspace:this.workspace,id});
  const response=await sqliteFetch(path,{method,body,headers:method==='POST'?{'Content-Type':'application/json'}:undefined});
  if(getStorageAccount()!==this.account)throw new Error('账户已变更，待提交队列已暂停');
  if(!response.ok){let reason='SQLite 未能保存待提交改动，请保留页面或导出草稿';try{reason=(await response.json()).error??reason;}catch{}throw new Error(reason);}
  return response;
 }
 async list(id:string):Promise<QueueSnapshot>{return (await this.request(id)).json();}
 async append(id:string,data:Uint8Array,snapshot?:Uint8Array){const encode=(bytes:Uint8Array)=>{let value='';for(let i=0;i<bytes.length;i+=8192)value+=String.fromCharCode(...bytes.subarray(i,i+8192));return btoa(value);};return (await this.request(id,'POST',JSON.stringify({update:encode(data),snapshot:snapshot?encode(snapshot):undefined}))).json() as Promise<{seq:number}>;}
 async acknowledge(id:string,seq:number){await this.request(id,'DELETE',JSON.stringify({seq}));}
 async fail(id:string,seq:number,error:string){await this.request(id,'PATCH',JSON.stringify({seq,error}));}
}
