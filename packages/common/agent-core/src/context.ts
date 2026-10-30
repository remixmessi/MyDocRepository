import type {Message,ToolSpec} from '@mydoc/agent-contracts';
export class BudgetError extends Error {}
export interface ContextPolicy {
 inputBudget:number;
 triggerRatio?:number;
 keepRecentMessages?:number;
 maxSummaryChars?:number;
 maxCompactionInputChars?:number;
 /** Must include protocol overhead; default deliberately counts UTF-8 bytes, not words. */
 estimate?:(messages:readonly Message[],tools:readonly ToolSpec[])=>number;
 compact?:(messages:readonly Message[],signal:AbortSignal)=>Promise<string>;
}
export function estimateContext(messages:readonly Message[],tools:readonly ToolSpec[]):number{
 return new TextEncoder().encode(JSON.stringify({messages,tools})).length+256;
}
/** Per-run context projection. The original journal is never rewritten. */
export class ContextManager {
 private summary='';
 private consumed=1;
 constructor(private readonly policy:ContextPolicy){
  if(!Number.isSafeInteger(policy.inputBudget)||policy.inputBudget<1||
   !Number.isFinite(policy.triggerRatio??0.8)||(policy.triggerRatio??0.8)<=0||(policy.triggerRatio??0.8)>1||
   !Number.isSafeInteger(policy.keepRecentMessages??4)||(policy.keepRecentMessages??4)<0||
   !Number.isSafeInteger(policy.maxCompactionInputChars??128_000)||(policy.maxCompactionInputChars??128_000)<1||
   !Number.isSafeInteger(policy.maxSummaryChars??8000)||(policy.maxSummaryChars??8000)<1)throw new Error('Invalid context policy');
 }
 async prepare(messages:readonly Message[],tools:readonly ToolSpec[],signal:AbortSignal,prefix:readonly Message[]=[]):Promise<Message[]>{
  const project=():Message[]=>[...prefix,...messages.slice(0,1),...(this.summary?[{role:'summary' as const,text:this.summary}]:[]),...messages.slice(this.consumed)];
  const count=(value:readonly Message[])=>{
   const size=(this.policy.estimate??estimateContext)(value,tools);
   if(!Number.isFinite(size)||size<0)throw new Error('Invalid context estimate');return size;
  };
  let context=project();
  if(count(context)>this.policy.inputBudget*(this.policy.triggerRatio??0.8)&&this.policy.compact){
   let cut=Math.max(this.consumed,messages.length-(this.policy.keepRecentMessages??4));
   // Never retain an orphan tool result or split an assistant/tool exchange.
   while(cut>this.consumed&&messages[cut]?.role==='tool')cut--;
   if(cut>this.consumed){
    const prefix:Message[]=[...(this.summary?[{role:'summary',text:this.summary} as Message]:[]),...messages.slice(this.consumed,cut)];
    if(JSON.stringify(prefix).length>(this.policy.maxCompactionInputChars??128_000))throw new BudgetError('Compaction input budget exceeded');
    const summary=await this.policy.compact(structuredClone(prefix),signal);
    if(typeof summary!=='string'||!summary.trim()||summary.length>(this.policy.maxSummaryChars??8000))throw new Error('Invalid context summary');
    this.summary=summary;this.consumed=cut;context=project();
   }
  }
  if(count(context)>this.policy.inputBudget)throw new BudgetError('Context input budget exceeded');
  return structuredClone(context);
 }
}
