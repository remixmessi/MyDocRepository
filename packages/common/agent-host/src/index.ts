export {createModelProvider} from './model-provider';
export {CompatibleProvider} from './compatible-provider';
export type {CompatibleModelConfig} from './compatible-provider';
import {runAgent} from '@mydoc/agent-core';
import type {RunOptions,RunResult} from '@mydoc/agent-core';
/** In-memory admission only; A1 replaces this with persistent queue/lease ownership. */
export class AgentHost {
 private readonly active=new Map<string,{runId:string;controller:AbortController}>();
 async start(options:Omit<RunOptions,'signal'>):Promise<RunResult>{
  if(this.active.has(options.sessionId))throw new Error('Session already running');
  const controller=new AbortController();this.active.set(options.sessionId,{runId:options.runId,controller});
  try{return await runAgent({...options,signal:controller.signal});}
  finally{this.active.delete(options.sessionId);}
 }
 cancel(sessionId:string,runId:string):boolean{
  const run=this.active.get(sessionId);if(!run||run.runId!==runId)return false;
  run.controller.abort();return true;
 }
}
