import {MemoryResultPages} from './result-pages';
export {MemoryResultPages} from './result-pages';
import {SourceContext} from './sources';
import type {ContextAccessPort} from './sources';
export {SourceContext} from './sources';
export type {ContextAccessPort,SpaceContext,SourceExcerpt} from './sources';
import {BudgetError,ContextManager} from './context';
import type {ContextPolicy} from './context';
import {executeBatch} from './batch';
export {ContextManager,estimateContext} from './context';
export type {ContextPolicy} from './context';
import {AGENT_PROTOCOL_VERSION,DEFAULT_RUN_LIMITS,validateCall,validateLimits,validateTools,validateUsage} from '@mydoc/agent-contracts';
import type {Clock,EventPayload,EventPort,Message,ProviderPort,RunIdentity,RunLimits,RunStatus,TerminalStatus,ToolCall,ToolGateway,ToolSpec,InteractionPort,ResultPages,TokenUsage} from '@mydoc/agent-contracts';
export interface RunOptions extends RunIdentity {
 prompt:string;provider:ProviderPort;gateway:ToolGateway;tools:readonly ToolSpec[];
 events:EventPort;clock:Clock;resultPages?:ResultPages;context?:ContextPolicy;sourceContext?:ContextAccessPort;interaction?:InteractionPort;signal?:AbortSignal;limits?:Partial<RunLimits>;
}
export interface RunResult {status:TerminalStatus;messages:Message[];resultPages:ResultPages;usage:TokenUsage;usageReports:number}
class EventWriteError extends Error {}
/** Race uncooperative adapters as well as forwarding cancellation to cooperative ones. */
async function interruptible<T>(work:Promise<T>,signal:AbortSignal):Promise<T>{
 if(signal.aborted){void work.catch(()=>{});throw signal.reason;}
 let listener:()=>void=()=>{};
 const interrupted=new Promise<never>((_,reject)=>{listener=()=>reject(signal.reason);signal.addEventListener('abort',listener,{once:true});});
 try{return await Promise.race([work,interrupted]);}finally{signal.removeEventListener('abort',listener);}
}
const transitions:Record<RunStatus,readonly RunStatus[]>={
 queued:['preparing','interrupted'],preparing:['streaming','failed','cancelled','budget_exhausted'],
 streaming:['executing_tools','finalizing','failed','cancelled','budget_exhausted'],
 executing_tools:['streaming','waiting_input','waiting_approval','failed','cancelled','budget_exhausted'],
 waiting_input:['executing_tools','failed','cancelled','budget_exhausted'],
 waiting_approval:['executing_tools','failed','cancelled','budget_exhausted'],
 cancelling:['cancelled','failed','budget_exhausted'],interrupted:[],
 finalizing:['completed','failed','cancelled','budget_exhausted'],completed:[],cancelled:[],failed:[],budget_exhausted:[],
};
for(const state of ['preparing','streaming','executing_tools','waiting_input','waiting_approval','finalizing'] as const)transitions[state]=[...transitions[state],'cancelling','interrupted'];
export function canTransition(from:RunStatus,to:RunStatus){return transitions[from].includes(to);}
/** One run only. Persistence, authentication and per-session admission belong to the host. */
export async function runAgent(options:RunOptions):Promise<RunResult>{
 const limits={...DEFAULT_RUN_LIMITS,...options.limits};validateLimits(limits);
 if(typeof options.prompt!=='string'||!options.prompt.trim()||options.prompt.length>limits.maxInputChars)throw new Error('Invalid prompt');
 validateTools(options.tools);
 const context=new ContextManager(options.context??{inputBudget:128_000});
 const resultPages=options.resultPages??new MemoryResultPages();
 const usage:TokenUsage={inputTokens:0,outputTokens:0,totalTokens:0};
 const sources=options.sourceContext?new SourceContext(options.sourceContext):undefined;
 const names=new Set(options.tools.map(tool=>tool.name));
 if(names.size!==options.tools.length)throw new Error('Duplicate tool name');
 const controller=new AbortController(),signal=controller.signal;
 const cancel=()=>controller.abort(new Error('Cancelled'));
 options.signal?.addEventListener('abort',cancel,{once:true});if(options.signal?.aborted)cancel();
 const timer=setTimeout(()=>controller.abort(new BudgetError('Run deadline exceeded')),limits.timeoutMs);
 let seq=0,status:RunStatus='queued',outputChars=0,usageReports=0;
 let journal:Promise<void>=Promise.resolve();
 const messages:Message[]=[];
 let pending:ToolCall[]=[],text='',recorded=false;
 const started=new Set<string>();
 const emit=(payload:EventPayload)=>{
  // Serialize journal writes even when read tools finish concurrently. A failure poisons the queue.
  journal=journal.then(async()=>{
   try{await options.events.append({protocolVersion:AGENT_PROTOCOL_VERSION,sessionId:options.sessionId,runId:options.runId,seq:++seq,timestamp:options.clock.now(),payload});}
   catch(error){throw new EventWriteError(error instanceof Error?error.message:'Event persistence failed');}
  });
  return journal;
 };
 const move=async(next:RunStatus)=>{if(!canTransition(status,next))throw new Error(`Invalid transition ${status} → ${next}`);await emit({type:'status',status:next});status=next;};
 const check=()=>{if(signal.aborted)throw signal.reason;};
 const add=async(message:Message)=>{await emit({type:'message',message});messages.push(message);};
 const saveAssistant=async()=>{if(!recorded&&(text||pending.length)){await add({role:'assistant',text,calls:[...pending]});recorded=true;}};
 const answer=async(call:ToolCall,result:{status:'ok'|'error'|'denied'|'skipped';text:string})=>{
  const truncated=result.text.length>limits.maxToolResultChars;
  const artifactId=options.runId+':'+call.id;
  if(truncated)await resultPages.put(artifactId,result.text);
  const value=truncated?result.text.slice(0,limits.maxToolResultChars)+'\n[truncated; result='+artifactId+'; offset='+limits.maxToolResultChars+']':result.text;
  await add({role:'tool',callId:call.id,status:result.status,text:value});
  pending=pending.filter(item=>item.id!==call.id);
 };
 try{
  await move('preparing');check();await add({role:'user',text:options.prompt});
  const seen=new Set<string>();
  for(let round=0;round<limits.maxRounds;round++){
   check();if(usage.totalTokens>=limits.maxTotalTokens)throw new BudgetError('Total token budget exhausted');
   await move('streaming');pending=[];text='';recorded=false;
   let done=false,usageSeen=false;
   const prefix=sources?await interruptible(sources.prepare(options.tools,signal),signal):[];check();
   const modelContext=await interruptible(context.prepare(messages,options.tools,signal,prefix),signal);check();
   // Compaction may await another model; recheck authorization before sending retained content.
   if(sources&&JSON.stringify(await interruptible(sources.prepare(options.tools,signal),signal))!==JSON.stringify(prefix))throw new Error('Source context changed; restart run');
   check();
   const iterator=options.provider.stream(modelContext,structuredClone(options.tools),signal)[Symbol.asyncIterator]();
   try{
    while(true){
     check();const part=await interruptible(iterator.next(),signal);check();
     if(part.done)break;
     const event=part.value;
     if(event.type==='done'){done=true;break;}
     if(event.type==='usage'){
      if(usageSeen)throw new Error('Duplicate model usage');usageSeen=true;
      validateUsage(event.usage);usageReports++;
      usage.inputTokens+=event.usage.inputTokens;usage.outputTokens+=event.usage.outputTokens;usage.totalTokens+=event.usage.totalTokens;
      await emit({type:'usage',usage:{...usage}});
      if(usage.totalTokens>limits.maxTotalTokens)throw new BudgetError('Total token budget exceeded');
      continue;
     }
     if(event.type==='text'){
      if(typeof event.text!=='string')throw new Error('Invalid provider text');
      outputChars+=event.text.length;if(outputChars>limits.maxOutputChars)throw new BudgetError('Output budget exceeded');
      text+=event.text;await emit({type:'text',text:event.text});
     }else if(event.type==='tool'){
      validateCall(event.call);if(seen.has(event.call.id))throw new Error('Duplicate tool call id');
      if(pending.length>=limits.maxCallsPerRound)throw new BudgetError('Tool call budget exceeded');
      seen.add(event.call.id);pending.push(structuredClone(event.call));
     }else throw new Error('Unknown provider event');
    }
   }finally{
    // Never await an uncooperative stream's cleanup during cancellation.
    void Promise.resolve(iterator.return?.()).catch(()=>{});
   }
   await saveAssistant();if(!done)throw new Error('Provider stream ended without completion');check();
   if(!pending.length){await move('finalizing');check();await move('completed');return {status:'completed',messages,resultPages,usage,usageReports};}
   await move('executing_tools');
   // Admission and human waits occur before taking any read execution slots.
   const admitted:ToolCall[]=[];
   for(const call of [...pending]){
    check();
    if(!names.has(call.name)){await answer(call,{status:'denied',text:'Tool unavailable'});continue;}
    let admission=await interruptible(options.gateway.prepare?.(structuredClone(call),signal)??Promise.resolve({status:'ready' as const}),signal);check();
    if(admission.status==='wait'){
     const request=admission.request;
     if(!request.id||request.id.length>128||!request.prompt||request.prompt.length>8000||!['input','approval'].includes(request.kind))throw new Error('Invalid interaction request');
     if(!options.interaction){await answer(call,{status:'denied',text:'Interaction unavailable'});continue;}
     await move(request.kind==='input'?'waiting_input':'waiting_approval');
     await emit({type:'interaction',request:structuredClone(request)});
     const response=await interruptible(options.interaction.wait(structuredClone(request),signal),signal);check();
     if(typeof response!=='string'||response.length>8000)throw new Error('Invalid interaction response');
     admission=await interruptible(options.gateway.prepare!(structuredClone(call),signal,{requestId:request.id,answer:response}),signal);check();
     await move('executing_tools');
    }
    if(admission.status!=='ready'){await answer(call,{status:'denied',text:admission.status==='denied'?admission.reason:'Admission still requires interaction'});continue;}
    admitted.push(call);
   }
   await executeBatch(admitted,options.tools,limits.maxReadConcurrency,async(call)=>{
    check();
    let result;
    try{started.add(call.id);result=await interruptible(options.gateway.execute(structuredClone(call),signal),signal);}
    catch(error){check();result={status:'error' as const,text:error instanceof Error?error.message:'Tool failed'};}
    check();
    if(!result||typeof result.text!=='string'||!['ok','error','denied','skipped'].includes(result.status))throw new Error('Invalid tool result');
    await answer(call,result);
   },signal);
  }
  throw new BudgetError('Round budget exceeded');
 }catch(error){
  // A failed journal cannot safely record recovery; leave ownership recovery to the host.
  if(error instanceof EventWriteError)throw error;
  // Preserve accepted calls and pair every unexecuted call, even on a partial stream.
  await saveAssistant();
  for(const call of [...pending])await answer(call,{status:started.has(call.id)?'error':'skipped',text:started.has(call.id)?'Execution outcome unconfirmed; inspect operation receipt before retry':'Not executed: run stopped'});
  const reason=signal.aborted?signal.reason:error;
  const terminal=reason instanceof BudgetError?'budget_exhausted':signal.aborted?'cancelled':'failed';
  await emit({type:'error',message:reason instanceof Error?reason.message:'Agent failed'});
  if(terminal==='cancelled')await move('cancelling');
  await move(terminal);return {status:terminal,messages,resultPages,usage,usageReports};
 }finally{clearTimeout(timer);options.signal?.removeEventListener('abort',cancel);}
}
