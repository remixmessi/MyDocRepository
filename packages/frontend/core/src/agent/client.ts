import {localAccount} from '../local-account-ui';
import type {ModelApiType} from '@mydoc/agent-contracts';
import {projectPreferences} from '@mydoc/nbstore';
import {cloudServerUrl} from '../../../../common/nbstore/src/cloud';
import {getStorageAccount,getStorageGeneration} from '../../../../common/nbstore/src/desktop';
export class AgentRequestError extends Error {constructor(message:string,readonly status:number){super(message);}}
export interface EventRow {cursor:number;payload:string}
export interface AgentConnection {id:string;name:string;models:string[];enabled:boolean;apiType?:ModelApiType;endpoint:string;model:string;hasKey:boolean;keyAvailable:boolean;maxOutputTokens:number;contextWindow:number;supportsTools:boolean;includeUsage:boolean}
export interface AgentStatus {models?:{version:1;connections:AgentConnection[];defaultModel?:{connectionId:string;model:string}};generation?:string;enabled:boolean;managed:boolean;userConfig:boolean;secureStorage:boolean;secretStorageLabel?:string;disclosure:string;model?:{apiType?:ModelApiType;endpoint:string;model:string;hasKey:boolean;keyAvailable:boolean;maxOutputTokens?:number;contextWindow?:number;supportsTools?:boolean;includeUsage?:boolean}}
let mode:'local'|'server'='local';
const generations=new Map<string,string>();
export function setAgentMode(value:'local'|'server'){mode=value;}
export function getAgentMode(){return mode;}
export function getAgentIdentity(){return JSON.stringify([mode,mode==='local'?(window.mydocDesktop?projectPreferences.getItem('local-account-session')??'':localAccount?.id??''):getStorageAccount()??'',mode==='server'?cloudServerUrl():'']);}
function endpoint(){if(mode==='local')return '/api/local-agent';const server=cloudServerUrl();if(window.mydocDesktop)throw new Error('桌面团队 Agent 尚未接入，请在团队服务器网页中使用');if(server&&new URL(server).origin!==location.origin)throw new Error('跨服务器 Agent 尚未接入，请打开对应团队服务器页面');return server+'/api/agent';}
function identityHeaders():Record<string,string>{return mode==='local'?(localAccount?{'X-Mydoc-Account':localAccount.id}:{}):getStorageAccount()?{'X-Mydoc-Account':getStorageAccount()!}:{};}
export async function agentRequest<T>(path:string,input?:unknown,signal?:AbortSignal):Promise<T>{
 const account=getAgentIdentity();const base=endpoint();const headers={'Content-Type':'application/json',...identityHeaders(),...((generations.get(account)??getStorageGeneration())?{'X-Mydoc-Generation':(generations.get(account)??getStorageGeneration())!}:{})};
 const response=window.mydocDesktop?.agent?await window.mydocDesktop.agent({path,method:input===undefined?'GET':'POST',body:input===undefined?undefined:new TextEncoder().encode(JSON.stringify(input)),generation:generations.get(account)??getStorageGeneration()}).then(reply=>new Response(reply.body.slice().buffer,{status:reply.status})):await fetch(base+path,{method:input===undefined?'GET':'POST',headers,body:input===undefined?undefined:JSON.stringify(input),signal});
 if(account!==getAgentIdentity())throw new Error('账户已变更，请重新打开 Agent');
 let value;try{value=await response.json();}catch{throw new Error('当前服务尚未启用 Agent，请使用独立服务或桌面版');}
 if(!response.ok)throw new AgentRequestError(value.error??'Agent 请求失败',response.status);if(path==='/status'&&typeof value.generation==='string'){if(generations.size>100)generations.clear();generations.set(account,value.generation);}return value;
}
export async function subscribe(session:string,after:number,onRow:(row:EventRow)=>void,signal:AbortSignal){
 if(window.mydocDesktop){for(;;){if(signal.aborted)return;const rows=await agentRequest<EventRow[]>('/events?session='+encodeURIComponent(session)+'&after='+after,undefined,signal);if(signal.aborted)return;for(const row of rows){if(row.cursor>after){onRow(row);after=row.cursor;}}await new Promise<void>(resolve=>{const done=()=>{clearTimeout(timer);signal.removeEventListener('abort',done);resolve();};const timer=setTimeout(done,500);signal.addEventListener('abort',done,{once:true});});}}
 const account=getAgentIdentity();const response=await fetch(endpoint()+'/subscribe?session='+encodeURIComponent(session)+'&after='+after,{headers:identityHeaders(),signal});
 if(!response.ok)throw new AgentRequestError(response.status===401?'登录已失效，请重新登录后打开 Agent':'事件连接失败，请重试',response.status);
 if(!response.body||!response.headers.get('content-type')?.includes('text/event-stream'))throw new Error('事件连接失败，请重试');
 const reader=response.body.getReader(),decoder=new TextDecoder();let buffer='';
 try{for(;;){const part=await reader.read();if(part.done)throw new Error('连接已断开');if(account!==getAgentIdentity())throw new Error('账户已变更');buffer+=decoder.decode(part.value,{stream:true});if(buffer.length>1_000_000)throw new Error('事件超过大小限制');let end;while((end=buffer.indexOf('\n\n'))>=0){const frame=buffer.slice(0,end);buffer=buffer.slice(end+2);const data=frame.split('\n').find(line=>line.startsWith('data: '));if(data){const row=JSON.parse(data.slice(6)) as EventRow;if(row.cursor>after){onRow(row);after=row.cursor;}}}}}finally{await reader.cancel().catch(()=>{});reader.releaseLock();}
}
