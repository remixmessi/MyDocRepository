import {runAgent} from '@mydoc/agent-core';
import type {ProviderEvent} from '@mydoc/agent-contracts';
let next=0;const pending=new Map<number,{resolve:(value:any)=>void;reject:(error:Error)=>void}>();
const controller=new AbortController();
function rpc(method:string,args:unknown[]):Promise<any>{return new Promise((resolve,reject)=>{const id=++next;pending.set(id,{resolve,reject});process.send!({type:'rpc',id,method,args});});}
// Stream pull is RPC-based, so IPC queues remain bounded by one outstanding frame.
process.on('message',async(raw:any)=>{
 if(raw.type==='reply'){const request=pending.get(raw.id);if(request){pending.delete(raw.id);raw.error?request.reject(new Error(raw.error)):request.resolve(raw.value);}return;}
 if(raw.type==='cancel'){controller.abort();return;}
 if(raw.type!=='start')return;
 try{
  const result=await runAgent({...raw.options,signal:controller.signal,clock:{now:()=>Date.now()},
   sourceContext:raw.sources?{snapshot:()=>rpc('sourceSnapshot',[]),authorize:(ref:any)=>rpc('sourceAuthorize',[ref])}:undefined,
   interaction:raw.interaction?{wait:(request:any)=>rpc('interaction',[request])}:undefined,
   context:raw.options.context?{...raw.options.context,compact:raw.compact?(messages:any)=>rpc('compact',[messages]):undefined}:undefined,
   provider:{async *stream(messages:any,tools:any){await rpc('modelStart',[messages,tools]);while(true){const part=await rpc('modelNext',[]);if(part.done)return;yield part.value as ProviderEvent;}}},
   gateway:{prepare:raw.prepare?(...args:any[])=>rpc('prepare',args.slice(0,1).concat(args[2]??null)):undefined,execute:(call:any)=>rpc('execute',[call])},
   events:{append:(event:any)=>rpc('event',[event])},
   resultPages:{put:(id:string,text:string)=>rpc('resultPut',[id,text]),read:(...args:any[])=>rpc('resultRead',args)},
  });
  process.send!({type:'finished',status:result.status});
 }catch{process.send!({type:'failed'});}
});
process.on('disconnect',()=>process.exit(0));
