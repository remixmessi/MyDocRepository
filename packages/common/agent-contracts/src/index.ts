export * from './lifecycle';
import type {TokenUsage} from './lifecycle';
/** Portable contracts: no platform, storage or UI dependencies. */
export type JsonValue = null|boolean|number|string|JsonValue[]|{[key:string]:JsonValue};
/** Reject lossy serialization (functions, undefined, NaN, custom objects) at the boundary. */
export function validateJson(value:unknown,depth=0,seen=new Set<object>()):asserts value is JsonValue {
 if(depth>40)throw new Error('JSON nesting limit exceeded');
 if(value===null||typeof value==='string'||typeof value==='boolean')return;
 if(typeof value==='number'&&Number.isFinite(value))return;
 if(typeof value!=='object'||!value||seen.has(value))throw new Error('Invalid JSON value');
 if(!Array.isArray(value)&&Object.getPrototypeOf(value)!==Object.prototype&&Object.getPrototypeOf(value)!==null)throw new Error('Invalid JSON object');
 seen.add(value);
 for(const child of Array.isArray(value)?value:Object.values(value))validateJson(child,depth+1,seen);
 seen.delete(value);
}
export const AGENT_PROTOCOL_VERSION = 1;
export type RunStatus = 'queued'|'preparing'|'streaming'|'executing_tools'|'waiting_input'|'waiting_approval'|'cancelling'|'interrupted'|'finalizing'|'completed'|'cancelled'|'failed'|'budget_exhausted';
export type TerminalStatus = Extract<RunStatus,'completed'|'cancelled'|'failed'|'budget_exhausted'|'interrupted'>;
export type ToolResultStatus = 'ok'|'error'|'denied'|'skipped';
export interface ToolCall {id:string;name:string;arguments:unknown}
export type Message =
 | {role:'user';text:string}
 | {role:'summary';text:string}
 | {role:'system';text:string}
 | {role:'assistant';text:string;calls:ToolCall[]}
 | {role:'tool';callId:string;status:ToolResultStatus;text:string};
export interface ToolSpec {name:string;effect:'read'|'write';description:string;inputSchema?:JsonValue;outputSchema?:JsonValue}
export type ProviderEvent = {type:'text';text:string}|{type:'tool';call:ToolCall}|{type:'usage';usage:TokenUsage}|{type:'done'};
export interface ProviderPort {
 /** tool events contain complete arguments; partial transport JSON stays in the adapter. */
 stream(messages:readonly Message[],tools:readonly ToolSpec[],signal:AbortSignal):AsyncIterable<ProviderEvent>;
}
export interface InteractionRequest {id:string;kind:'input'|'approval';prompt:string}
export type ToolAdmission = {status:'ready'}|{status:'denied';reason:string}|{status:'wait';request:InteractionRequest};
export interface InteractionPort {wait(request:InteractionRequest,signal:AbortSignal):Promise<string>}
export interface SourceRef {serverId:string;workspaceId:string;resourceId:string;kind:string;anchor:string;revision:string;excerptHash:string}
export interface ChangeSet {id:string;accountScope:string;workspaceId:string;resourceId:string;kind:string;baseRevision:string;generation:string;operations:unknown[];preview:string;contentHash:string;status:'proposed'|'applied'|'discarded'|'conflict'}
export interface ToolGateway {
 /** Re-run authoritative admission with the user's answer; wait does not itself grant permission. */
 prepare?(call:ToolCall,signal:AbortSignal,response?:{requestId:string;answer:string}):Promise<ToolAdmission>;
 /** Implementations validate arguments, authorize the principal and fence stale runs. */
 execute(call:ToolCall,signal:AbortSignal):Promise<{status:ToolResultStatus;text:string}>;
}
export interface RunIdentity {sessionId:string;runId:string}
export type EventPayload = {type:'status';status:RunStatus}|{type:'text';text:string}|{type:'message';message:Message}|{type:'error';message:string}|{type:'interaction';request:InteractionRequest}|{type:'usage';usage:TokenUsage};
export type AgentEvent = RunIdentity & {protocolVersion:typeof AGENT_PROTOCOL_VERSION;seq:number;timestamp:number;payload:EventPayload};
export interface EventPort {/** Must resolve after accepting the event; reject to stop execution. */ append(event:AgentEvent):Promise<void>}
export interface Clock {now():number}
export interface RunLimits {maxRounds:number;timeoutMs:number;maxOutputChars:number;maxToolResultChars:number;maxCallsPerRound:number;maxReadConcurrency:number;maxInputChars:number;maxTotalTokens:number}
export const DEFAULT_RUN_LIMITS:Readonly<RunLimits> = Object.freeze({maxRounds:20,timeoutMs:600_000,maxOutputChars:100_000,maxToolResultChars:8_000,maxCallsPerRound:32,maxReadConcurrency:4,maxInputChars:100_000,maxTotalTokens:128_000});
export function validateLimits(limits:RunLimits):void {
 for(const [key,value] of Object.entries(limits))if(!Number.isSafeInteger(value)||value<=0)throw new Error(`Invalid limit: ${key}`);
}
export function validateCall(value:unknown):asserts value is ToolCall {
 if(!value||typeof value!=='object')throw new Error('Invalid tool call');
 const call=value as Partial<ToolCall>;
 if(typeof call.id!=='string'||!call.id.length||call.id.length>128||typeof call.name!=='string'||!/^[a-zA-Z][a-zA-Z0-9_.-]{0,63}$/.test(call.name))throw new Error('Invalid tool identity');
 validateJson(call.arguments);
 const json=JSON.stringify(call.arguments);
 if(json===undefined||json.length>65_536)throw new Error('Invalid tool arguments');
}

export function validateTools(tools:readonly ToolSpec[]):void {
 if(!Array.isArray(tools)||tools.length>128)throw new Error('Invalid tool catalog');
 const names=new Set<string>();
 for(const tool of tools){
  if(!tool||!['read','write'].includes(tool.effect)||typeof tool.description!=='string'||tool.description.length>8000)throw new Error('Invalid tool specification');
  validateCall({id:'spec',name:tool.name,arguments:{}});
  if(names.has(tool.name))throw new Error('Duplicate tool name');names.add(tool.name);
  if(tool.inputSchema!==undefined)validateJson(tool.inputSchema);
  if(tool.outputSchema!==undefined)validateJson(tool.outputSchema);
 }
 if(JSON.stringify(tools).length>128_000)throw new Error('Tool catalog size limit exceeded');
}
