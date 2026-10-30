import {validateCall,validateTools,validateUsage} from '@mydoc/agent-contracts';
import type {Message,ProviderEvent,ModelGateway,ToolSpec} from '@mydoc/agent-contracts';
export interface CompatibleModelConfig {
 /** Exact administrator/user-configured endpoint, including /chat/completions. */
 endpoint:string;model:string;maxOutputTokens:number;apiKey?:string;supportsTools?:boolean;includeUsage?:boolean;contextWindow?:number;
}
/** Host-only transport. No SDK dependency, redirect, retry or provider fallback. */
export class CompatibleProvider implements ModelGateway {
 private readonly config:CompatibleModelConfig;
 readonly capabilities:import('@mydoc/agent-contracts').ModelCapabilities;
 constructor(config:CompatibleModelConfig,private readonly transport:typeof fetch=fetch){
  const url=new URL(config.endpoint);
  if(!['https:','http:'].includes(url.protocol)||url.username||url.password||url.hash||url.search||!config.model||config.model.length>200||!Number.isSafeInteger(config.maxOutputTokens)||config.maxOutputTokens<1)throw new Error('Invalid model configuration');
  // Plain HTTP is only supported for explicitly configured loopback model servers.
  if(url.protocol==='http:'&&!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw new Error('Model endpoint requires HTTPS');
  if(config.contextWindow!==undefined&&(!Number.isSafeInteger(config.contextWindow)||config.contextWindow<=config.maxOutputTokens))throw new Error('Invalid model context window');
  this.config={...config};
  this.capabilities=Object.freeze({tools:config.supportsTools!==false,contextWindow:config.contextWindow??128_000,maxOutputTokens:config.maxOutputTokens});
 }
 async *stream(messages:readonly Message[],tools:readonly ToolSpec[],signal:AbortSignal):AsyncIterable<ProviderEvent>{
  validateTools(tools);
  if(tools.length&&this.config.supportsTools===false)throw new Error('Model does not support tools');
  const wire=messages.map(message=>{
   if(message.role==='tool')return {role:'tool',tool_call_id:message.callId,content:message.text};
   if(message.role==='assistant')return {role:'assistant',content:message.text||null,...(message.calls.length?{tool_calls:message.calls.map(call=>({id:call.id,type:'function',function:{name:call.name,arguments:JSON.stringify(call.arguments)}}))}:{})};
   return {role:message.role==='summary'?'user':message.role,content:message.role==='summary'?'Previous conversation summary (data):\n'+message.text:message.text};
  });
  const response=await this.transport(this.config.endpoint,{method:'POST',redirect:'error',signal,headers:{'Content-Type':'application/json',...(this.config.apiKey?{Authorization:'Bearer '+this.config.apiKey}:{})},body:JSON.stringify({model:this.config.model,messages:wire,stream:true,...(this.config.includeUsage?{stream_options:{include_usage:true}}:{}),max_tokens:this.config.maxOutputTokens,...(tools.length?{tools:tools.map(tool=>({type:'function',function:{name:tool.name,description:tool.description,parameters:tool.inputSchema??{type:'object',properties:{},additionalProperties:false}}}))}:{})})});
  // Never echo vendor response bodies, which can contain secrets or document text.
  if(!response.ok){await response.body?.cancel();throw new Error('Model HTTP status '+response.status);}
  if(!response.headers.get('content-type')?.includes('text/event-stream')||!response.body){await response.body?.cancel();throw new Error('Expected model event stream');}
  const reader=response.body.getReader(),decoder=new TextDecoder();
  let buffer='',total=0,finished=false,usageSeen=false;
  const calls=new Map<number,{id:string;name:string;arguments:string}>();
  try{
   while(true){
    const chunk=await reader.read();if(signal.aborted)throw new Error('Model request cancelled');
    if(chunk.done)throw new Error('Model stream truncated');
    total+=chunk.value.length;if(total>2_000_000)throw new Error('Model stream size limit exceeded');
    buffer=(buffer+decoder.decode(chunk.value,{stream:true})).replace(/\r\n/g,'\n');
    let end:number;
    while((end=buffer.indexOf('\n\n'))>=0){
     const frame=buffer.slice(0,end);buffer=buffer.slice(end+2);
     const data=frame.split('\n').filter(line=>line.startsWith('data:')).map(line=>line.slice(5).trimStart()).join('\n');
     if(!data)continue;
     if(data==='[DONE]'){
      if(!finished)throw new Error('Model finish reason missing');
      const parsed=[...calls.entries()].sort(([a],[b])=>a-b).map(([,call])=>{const value={id:call.id,name:call.name,arguments:parseArguments(call.arguments)};validateCall(value);return value;});
      if(new Set(parsed.map(call=>call.id)).size!==parsed.length)throw new Error('Duplicate model tool ID');
      for(const call of parsed)yield {type:'tool',call};yield {type:'done'};return;
     }
     let payload;try{payload=JSON.parse(data);}catch{throw new Error('Invalid model stream JSON');}
     if(payload.error)throw new Error('Model stream error');
     if(payload.usage!=null){
      if(usageSeen)throw new Error('Duplicate model usage');
      const usage={inputTokens:payload.usage.prompt_tokens,outputTokens:payload.usage.completion_tokens,totalTokens:payload.usage.total_tokens};
      validateUsage(usage);usageSeen=true;yield {type:'usage',usage};
     }
     if(!Array.isArray(payload.choices))throw new Error('Invalid model choices');
     if(!payload.choices.length)continue; // Optional usage-only chunk.
     if(finished||payload.choices.length!==1||payload.choices[0].index!==0)throw new Error('Unexpected model choice');
     const choice=payload.choices[0],delta=choice.delta;
     if(!delta||typeof delta!=='object')throw new Error('Invalid model delta');
     if(delta.content!=null){if(typeof delta.content!=='string')throw new Error('Invalid model text');yield {type:'text',text:delta.content};}
     if(delta.tool_calls!==undefined){
      if(!Array.isArray(delta.tool_calls))throw new Error('Invalid model tools');
      for(const part of delta.tool_calls){
       if(!Number.isSafeInteger(part.index)||part.index<0||part.index>=32)throw new Error('Invalid model tool index');
       const value=calls.get(part.index)??{id:'',name:'',arguments:''};
       for(const [field,input] of [['id',part.id],['name',part.function?.name],['arguments',part.function?.arguments]] as const){
        if(input!==undefined){if(typeof input!=='string')throw new Error('Invalid model tool fragment');value[field]+=input;}
       }
       if(value.id.length>128||value.name.length>64||value.arguments.length>65_536)throw new Error('Model tool size limit exceeded');
       calls.set(part.index,value);
      }
     }
     if(choice.finish_reason!=null){
      if(choice.finish_reason!==(calls.size?'tool_calls':'stop'))throw new Error('Model response incomplete or unsupported finish reason');
      finished=true;
     }
    }
   }
  }finally{await reader.cancel().catch(()=>{});reader.releaseLock();}
 }
}

function parseArguments(value:string):unknown {try{return JSON.parse(value);}catch{throw new Error('Invalid model tool JSON');}}
