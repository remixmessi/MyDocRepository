import {runAgent} from '@mydoc/agent-core';
import type {RunOptions,RunResult} from '@mydoc/agent-core';
import type {AgentRun} from '@mydoc/agent-contracts';
import {SQLiteAgentStore} from './sqlite-store';
export interface RuntimePort {run(options:RunOptions):Promise<RunResult>}
/** Trusted account host. Admission is persisted before execution; never auto-replays interrupted writes. */
export class DurableAgentHost {
 private closing=false;
 private active=new Map<string,{controller:AbortController;done:Promise<void>}>();
 constructor(private readonly store:SQLiteAgentStore,private readonly runtime:RuntimePort={run:runAgent},private readonly retainRuns=50){}
 enqueue(session:string,id:string,key:string,input:string){if(this.closing)throw new Error('Host closing');return this.store.enqueue(session,id,key,input);}
 async runNext(session:string,configure:(lease:AgentRun)=>Omit<RunOptions,'sessionId'|'runId'|'prompt'|'events'|'signal'>):Promise<boolean>{
  if(this.closing)return false;
  const next=this.store.queued(session)[0] as {id:string;input:string}|undefined;if(!next)return false;
  const lease=this.store.claim(next.id),controller=new AbortController();
  let finish!:()=>void;const done=new Promise<void>(resolve=>{finish=resolve;});this.active.set(lease.runId,{controller,done});
  try{
   const options=configure(lease),gateway=options.gateway,provider=options.provider;
   const check=()=>{this.store.assertLease(lease);if(controller.signal.aborted)throw new Error('Run cancelled');};
   await this.runtime.run({...options,sessionId:session,runId:lease.runId,prompt:next.input,signal:controller.signal,events:this.store.events(lease),provider:{async *stream(messages,tools,signal){check();for await(const event of provider.stream(messages,tools,signal)){check();yield event;}}},gateway:{
    prepare:gateway.prepare?async(...args)=>{check();const value=await gateway.prepare!(...args);check();return value;}:undefined,
    execute:async(...args)=>{check();const value=await gateway.execute(...args);check();return value;},
   }});
   // A runtime must commit a terminal event before reporting success.
   this.store.interrupt(lease);
  }catch(error){this.store.interrupt(lease);throw error;}
  finally{
   // Retention is housekeeping: a failure here must not mask the run's own outcome.
   try{this.store.compact(lease.runId);this.store.prune(session,this.retainRuns);}catch{/* retried after the next run */}
   this.active.delete(lease.runId);finish();
  }
  return true;
 }
 cancel(id:string){const active=this.active.get(id);if(active){active.controller.abort();return true;}return this.store.cancelQueued(id);}
 /** Caller first excludes every other owner of this account; normal DB opens do not recover. */
 recover(){if(this.active.size)throw new Error('Host still active');this.store.recoverActive();}
 async close(){this.closing=true;for(const run of this.active.values())run.controller.abort();await Promise.all([...this.active.values()].map(run=>run.done));}
}
