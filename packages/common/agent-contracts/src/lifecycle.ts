import {validateJson,validateCall} from './index';
import type {AgentEvent,RunIdentity,RunStatus,SourceRef,ChangeSet,TerminalStatus,ProviderPort} from './index';
export type AgentErrorCode='invalid_request'|'provider_error'|'permission_denied'|'conflict'|'cancelled'|'interrupted'|'budget_exhausted'|'storage_error';
export interface TokenUsage {inputTokens:number;outputTokens:number;totalTokens:number}
export interface AgentSession {id:string;ownerId:string;workspaceId:string;createdAt:number}
export interface AgentRun extends RunIdentity {status:RunStatus;generation:number;requestKey:string}
export interface AgentStep extends RunIdentity {id:string;ordinal:number;kind:'model'|'tools'}
export interface IdPort {next():string}
/** Persistence belongs to A1; recovery must fence the prior generation before scheduling. */
export interface RunStore {load(runId:string):Promise<AgentRun|null>;save(run:AgentRun,expectedGeneration:number):Promise<void>}
export interface ResultPage {text:string;nextOffset:number|null}
export interface ResultPages {put(id:string,text:string):Promise<void>;read(id:string,offset:number,limit:number):Promise<ResultPage>}
export const TERMINAL_STATUSES:readonly TerminalStatus[]=['completed','cancelled','failed','budget_exhausted','interrupted'];
/** Prefix of a started call whose outcome is unknown; hosts query domain receipts before any retry. */
export const UNCONFIRMED_RESULT='Execution outcome unconfirmed';
export function recoverRun(run:AgentRun):AgentRun {
 validateRun(run);
 return TERMINAL_STATUSES.includes(run.status as TerminalStatus)?{...run}:{...run,status:'interrupted',generation:run.generation+1};
}
export function validateUsage(value:unknown):asserts value is TokenUsage {
 const u=value as TokenUsage;
 if(!u||![u.inputTokens,u.outputTokens,u.totalTokens].every(v=>Number.isSafeInteger(v)&&v>=0)||u.totalTokens!==u.inputTokens+u.outputTokens)throw new Error('Invalid model usage');
}
function bounded(value:unknown,max=200):asserts value is string {if(typeof value!=='string'||!value.length||value.length>max)throw new Error('Invalid bounded string');}
export function validateRun(value:unknown):asserts value is AgentRun {
 const run=value as AgentRun;if(!run)throw new Error('Invalid run');
 bounded(run.sessionId);bounded(run.runId);bounded(run.requestKey);
 if(!Number.isSafeInteger(run.generation)||run.generation<0||!['queued','preparing','streaming','executing_tools','waiting_input','waiting_approval','cancelling','finalizing',...TERMINAL_STATUSES].includes(run.status))throw new Error('Invalid run state');
}
export function validateSourceRef(value:unknown):asserts value is SourceRef {
 const ref=value as SourceRef;if(!ref)throw new Error('Invalid source reference');
 for(const key of ['serverId','workspaceId','resourceId','kind','anchor','revision','excerptHash'] as const)bounded(ref[key],1000);
}
export function validateChangeSet(value:unknown):asserts value is ChangeSet {
 validateJson(value);const change=value as unknown as ChangeSet;
 for(const key of ['id','accountScope','workspaceId','resourceId','kind','baseRevision','generation','contentHash'] as const)bounded(change[key]);
 if(!Array.isArray(change.operations)||change.operations.length>500||typeof change.preview!=='string'||change.preview.length>100_000||JSON.stringify(value).length>256_000||!['proposed','applied','discarded','conflict'].includes(change.status))throw new Error('Invalid changeset');
}
export function validateEvent(value:unknown):asserts value is AgentEvent {
 validateJson(value);const event=value as unknown as AgentEvent;
 bounded(event.runId);bounded(event.sessionId);
 if(event.protocolVersion!==1||!Number.isSafeInteger(event.seq)||event.seq<1||!Number.isFinite(event.timestamp)||JSON.stringify(event).length>256_000||!event.payload||!['status','text','message','error','interaction','usage'].includes(event.payload.type))throw new Error('Invalid event envelope');
 const payload=event.payload;
 if(payload.type==='status')validateRun({sessionId:event.sessionId,runId:event.runId,requestKey:'event',generation:0,status:payload.status});
 if(payload.type==='usage')validateUsage(payload.usage);
 if(payload.type==='text'||payload.type==='error'){if(typeof (payload.type==='text'?payload.text:payload.message)!=='string')throw new Error('Invalid event text');}
 if(payload.type==='interaction'){bounded(payload.request?.id);bounded(payload.request?.prompt,8000);if(!['input','approval'].includes(payload.request.kind))throw new Error('Invalid interaction');}
 if(payload.type==='message'){
  const message=payload.message;
  if(!message||!['system','summary','user','assistant','tool'].includes(message.role)||typeof message.text!=='string')throw new Error('Invalid message');
  if(message.role==='assistant'){if(!Array.isArray(message.calls)||message.calls.length>32)throw new Error('Invalid message calls');for(const call of message.calls)validateCall(call);}
  if(message.role==='tool'){bounded(message.callId);if(!['ok','denied','error','skipped'].includes(message.status))throw new Error('Invalid result status');}
 }
}

export interface ModelCapabilities {tools:boolean;contextWindow:number;maxOutputTokens:number}
export interface ModelGateway extends ProviderPort {readonly capabilities:ModelCapabilities}
export function validateSession(value:unknown):asserts value is AgentSession {
 const session=value as AgentSession;if(!session)throw new Error('Invalid session');
 bounded(session.id);bounded(session.ownerId);bounded(session.workspaceId);
 if(!Number.isFinite(session.createdAt)||session.createdAt<0)throw new Error('Invalid session time');
}
export function validateStep(value:unknown):asserts value is AgentStep {
 const step=value as AgentStep;if(!step)throw new Error('Invalid step');
 bounded(step.id);bounded(step.runId);bounded(step.sessionId);
 if(!Number.isSafeInteger(step.ordinal)||step.ordinal<0||!['model','tools'].includes(step.kind))throw new Error('Invalid step');
}
