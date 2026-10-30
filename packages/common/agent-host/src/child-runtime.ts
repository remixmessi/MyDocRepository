import {MemoryResultPages} from '@mydoc/agent-core';
import {validateEvent,TERMINAL_STATUSES,DEFAULT_RUN_LIMITS} from '@mydoc/agent-contracts';
import type {Message,TerminalStatus,TokenUsage} from '@mydoc/agent-contracts';
import {fork} from 'node:child_process';
import type {RunOptions,RunResult} from '@mydoc/agent-core';
import type {ProviderEvent} from '@mydoc/agent-contracts';
import type {RuntimePort} from './durable-host';
/** Trusted bundled entry path only. Fault isolation, not an OS sandbox. No ambient secrets inherited. */
export class ChildAgentRuntime implements RuntimePort {
 constructor(private readonly entry:string,private readonly cancelGraceMs=2000){}
 run(options:RunOptions):Promise<RunResult>{
  if(options.context?.estimate)throw new Error('Synchronous custom estimator cannot cross worker boundary');
  return new Promise((resolve,reject)=>{
   const child=fork(this.entry,[],{stdio:['ignore','ignore','ignore','ipc'],execArgv:[],env:{NODE_ENV:'production',ELECTRON_RUN_AS_NODE:'1'}});
   const resultPages=options.resultPages??new MemoryResultPages();const messages:Message[]=[];let status:TerminalStatus|undefined,usage:TokenUsage={inputTokens:0,outputTokens:0,totalTokens:0},usageReports=0;
   let closed=false,watchdog:ReturnType<typeof setTimeout>|undefined,timer:ReturnType<typeof setTimeout>|undefined,model:AsyncIterator<ProviderEvent>|undefined;
   const cancel=()=>{if(child.connected)child.send({type:'cancel'});timer??=setTimeout(()=>child.kill('SIGKILL'),this.cancelGraceMs);};
   const finish=(error?:Error)=>{if(closed)return;closed=true;clearTimeout(timer);clearTimeout(watchdog);options.signal?.removeEventListener('abort',cancel);void Promise.resolve(model?.return?.()).catch(()=>{});child.kill();error?reject(error):status?resolve({status,messages,usage,usageReports,resultPages}):reject(new Error('Worker missing committed terminal'));};
   watchdog=setTimeout(()=>finish(new Error('Agent worker deadline exceeded')),(options.limits?.timeoutMs??DEFAULT_RUN_LIMITS.timeoutMs)+this.cancelGraceMs);
   child.on('error',()=>finish(new Error('Agent worker failed')));
   child.on('exit',()=>finish(new Error('Agent worker exited')));
   child.on('message',async(raw:any)=>{
    if(closed)return;
    if(raw.type==='finished'){finish();return;}
    if(raw.type==='failed'){finish(new Error('Agent worker execution failed'));return;}
    if(raw.type!=='rpc')return;
    try{
     let value:unknown;
     if(raw.method==='event'){
      const event=raw.args[0];validateEvent(event);await options.events.append(event);
      if(event.payload.type==='message')messages.push(event.payload.message);
      if(event.payload.type==='usage'){usage=event.payload.usage;usageReports++;}
      if(event.payload.type==='status'&&TERMINAL_STATUSES.includes(event.payload.status as TerminalStatus))status=event.payload.status as TerminalStatus;
     }
     else{
      if(options.signal?.aborted)throw new Error('Agent worker cancelled');
      if(raw.method==='sourceSnapshot')value=await options.sourceContext!.snapshot(options.signal??new AbortController().signal);
      else if(raw.method==='sourceAuthorize')value=await options.sourceContext!.authorize(raw.args[0],options.signal??new AbortController().signal);
      else if(raw.method==='interaction')value=await options.interaction!.wait(raw.args[0],options.signal??new AbortController().signal);
      else if(raw.method==='compact')value=await options.context!.compact!(raw.args[0],options.signal??new AbortController().signal);
      else if(raw.method==='resultPut')await resultPages.put(raw.args[0],raw.args[1]);
      else if(raw.method==='resultRead')value=await resultPages.read(raw.args[0],raw.args[1],raw.args[2]);
      else if(raw.method==='modelStart')model=options.provider.stream(raw.args[0],raw.args[1],options.signal??new AbortController().signal)[Symbol.asyncIterator]();
      else if(raw.method==='modelNext'){if(!model)throw new Error('No model stream');value=await model.next();}
      else if(raw.method==='prepare')value=await options.gateway.prepare!(raw.args[0],options.signal??new AbortController().signal,raw.args[1]??undefined);
      else if(raw.method==='execute')value=await options.gateway.execute(raw.args[0],options.signal??new AbortController().signal);
      else throw new Error('Unknown worker method');
     }
     if(!closed&&child.connected)child.send({type:'reply',id:raw.id,value});
    }catch{if(!closed&&child.connected)child.send({type:'reply',id:raw.id,error:'Host operation failed'});}
   });
   options.signal?.addEventListener('abort',cancel,{once:true});
   child.send({type:'start',prepare:!!options.gateway.prepare,sources:!!options.sourceContext,interaction:!!options.interaction,compact:!!options.context?.compact,options:{sessionId:options.sessionId,runId:options.runId,prompt:options.prompt,tools:options.tools,limits:options.limits,context:options.context?{...options.context,compact:undefined,estimate:undefined}:undefined}});
   if(options.signal?.aborted)cancel();
  });
 }
}
