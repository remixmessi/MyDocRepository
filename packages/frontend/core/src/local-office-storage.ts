import {getStorageAccount} from '../../../common/nbstore/src/desktop';
import {packOfficeUpdate} from '../../../common/nbstore/src/office-generation';
import {Y} from '@mydoc/crdt';
import type {DocumentStorage} from '@mydoc/nbstore';
import {cloudRequest,CLOUD_ORIGIN,officeDocumentGeneration} from '../../../common/nbstore/src/cloud';
/** Uses the authenticated local SQLite boundary, never a configured team server. */
export class LocalOfficeStorage implements DocumentStorage {
 private readonly account=getStorageAccount();
 constructor(private workspace:string){}
 private check(){if(this.account!==getStorageAccount())throw new Error('账户已变更，请保留草稿并重新打开文件');}
 private path(id:string){return '/document?'+new URLSearchParams({workspace:this.workspace,id});}
 async load(id:string){this.check();const response=await cloudRequest(this.path(id));this.check();return response.status===204?null:new Uint8Array(await response.arrayBuffer());}
 async append(id:string,update:Uint8Array){this.check();await cloudRequest(this.path(id),{method:'POST',body:packOfficeUpdate(update,officeDocumentGeneration(this.workspace,id))});}
 async remove(id:string){this.check();await cloudRequest('/office-resource?'+new URLSearchParams({workspace:this.workspace,id}),{method:'DELETE'});}
 connect(id:string,doc:Y.Doc,onError:(error:unknown)=>void){let disposed=false,running=false;const refresh=async()=>{if(disposed||running)return;running=true;try{const update=await this.load(id);if(!disposed){if(update)Y.applyUpdate(doc,update,CLOUD_ORIGIN);onError(undefined);}}catch(error){if(!disposed)onError(error);}finally{running=false;}};const timer=setInterval(()=>void refresh(),1500);return()=>{disposed=true;clearInterval(timer);};}
 close(){}
}
