import {validateCall,validateTools,validateUsage} from '@mydoc/agent-contracts';
import type {Message,ModelGateway,ProviderEvent,ToolCall,ToolSpec} from '@mydoc/agent-contracts';
import {CompatibleProvider,type CompatibleModelConfig} from './compatible-provider';

/** Host-only native Responses / Messages transport, using the same endpoint restrictions. */
export class NativeProvider implements ModelGateway {
 readonly capabilities;
 constructor(private readonly config:CompatibleModelConfig,private readonly apiType:'openai-responses'|'anthropic-messages',private readonly transport:typeof fetch=fetch){
  this.capabilities=new CompatibleProvider(config).capabilities;
 }
 async *stream(messages:readonly Message[],tools:readonly ToolSpec[],signal:AbortSignal):AsyncIterable<ProviderEvent>{
  validateTools(tools);
  if(tools.length&&this.config.supportsTools===false)throw new Error('Model does not support tools');
  const anthropic=this.apiType==='anthropic-messages';
  const response=await this.transport(this.config.endpoint,{method:'POST',redirect:'error',signal,headers:{'Content-Type':'application/json',...(anthropic?{'anthropic-version':'2023-06-01',...(this.config.apiKey?{'x-api-key':this.config.apiKey}:{})}:(this.config.apiKey?{Authorization:'Bearer '+this.config.apiKey}:{}))},body:JSON.stringify(anthropic?this.messagesBody(messages,tools):this.responsesBody(messages,tools))});
  if(!response.ok){await response.body?.cancel();throw new Error('Model HTTP status '+response.status);}
  if(!response.body||!response.headers.get('content-type')?.includes('text/event-stream')){await response.body?.cancel();throw new Error('Expected model event stream');}
  if(anthropic)yield* this.readMessages(response.body,signal);else yield* this.readResponses(response.body,signal);
 }
 private responsesBody(messages:readonly Message[],tools:readonly ToolSpec[]){
  const input:unknown[]=[];
  for(const message of messages){
   if(message.role==='tool')input.push({type:'function_call_output',call_id:message.callId,output:message.text});
   else if(message.role==='assistant'){
    if(message.text)input.push({role:'assistant',content:message.text});
    for(const call of message.calls)input.push({type:'function_call',call_id:call.id,name:call.name,arguments:JSON.stringify(call.arguments)});
   }else input.push({role:message.role==='summary'?'user':message.role,content:summaryText(message)});
  }
  return {model:this.config.model,input,stream:true,store:false,max_output_tokens:this.config.maxOutputTokens,...(tools.length?{tools:tools.map(tool=>({type:'function',name:tool.name,description:tool.description,parameters:tool.inputSchema??{type:'object',properties:{},additionalProperties:false},strict:false}))}:{})};
 }
 private messagesBody(messages:readonly Message[],tools:readonly ToolSpec[]){
  const system:string[]=[],wire:{role:'user'|'assistant';content:unknown[]}[]=[];
  const append=(role:'user'|'assistant',blocks:unknown[])=>{if(!blocks.length)return;const last=wire.at(-1);if(last?.role===role)last.content.push(...blocks);else wire.push({role,content:blocks});};
  for(const message of messages){
   if(message.role==='system'){system.push(message.text);continue;}
   if(message.role==='tool')append('user',[{type:'tool_result',tool_use_id:message.callId,content:message.text,is_error:message.status!=='ok'}]);
   else if(message.role==='assistant')append('assistant',[...(message.text?[{type:'text',text:message.text}]:[]),...message.calls.map(call=>({type:'tool_use',id:call.id,name:call.name,input:call.arguments}))]);
   else append('user',[{type:'text',text:summaryText(message)}]);
  }
  return {model:this.config.model,messages:wire,...(system.length?{system:system.join('\n\n')}:{}),stream:true,max_tokens:this.config.maxOutputTokens,...(tools.length?{tools:tools.map(tool=>({name:tool.name,description:tool.description,input_schema:tool.inputSchema??{type:'object',properties:{},additionalProperties:false}}))}:{})};
 }
 private async *readResponses(body:ReadableStream<Uint8Array>,signal:AbortSignal):AsyncIterable<ProviderEvent>{
  for await(const event of frames(body,signal)){
   if(event.type==='error'||event.type==='response.failed'||event.type==='response.incomplete')throw new Error('Model response incomplete or failed');
   if(event.type==='response.output_text.delta'){
    if(typeof event.delta!=='string')throw new Error('Invalid model text');
    yield {type:'text',text:event.delta};
   }
   if(event.type==='response.completed'){
    const response=event.response;
    if(response?.status!=='completed'||!Array.isArray(response.output))throw new Error('Model response incomplete');
    const calls:ToolCall[]=[];
    for(const item of response.output){
     if(item.type==='function_call')calls.push(checkedCall(item.call_id,item.name,parseArguments(item.arguments)));
     else if(!['message','reasoning'].includes(item.type))throw new Error('Unsupported model output');
    }
    uniqueCalls(calls);
    if(response.usage){const usage={inputTokens:response.usage.input_tokens,outputTokens:response.usage.output_tokens,totalTokens:response.usage.total_tokens};validateUsage(usage);yield {type:'usage',usage};}
    for(const call of calls)yield {type:'tool',call};
    yield {type:'done'};return;
   }
  }
  throw new Error('Model stream truncated');
 }
 private async *readMessages(body:ReadableStream<Uint8Array>,signal:AbortSignal):AsyncIterable<ProviderEvent>{
  const blocks=new Map<number,{type:string;id?:string;name?:string;json:string;input?:unknown;stopped:boolean}>();
  let started=false,finished=false,inputTokens=0,outputTokens=0,stopReason='';
  for await(const event of frames(body,signal)){
   if(event.type==='error')throw new Error('Model stream error');
   if(event.type==='message_start'){
    if(started)throw new Error('Duplicate model message');started=true;
    const usage=event.message?.usage;
    if(!usage)throw new Error('Missing model usage');
    inputTokens=usage.input_tokens+(usage.cache_creation_input_tokens??0)+(usage.cache_read_input_tokens??0);outputTokens=usage.output_tokens;
    validateUsage({inputTokens,outputTokens,totalTokens:inputTokens+outputTokens});
   }else if(event.type==='content_block_start'){
    if(!started||finished||!Number.isSafeInteger(event.index)||event.index<0||event.index>=128||blocks.has(event.index))throw new Error('Invalid model block');
    const block=event.content_block;
    if(!block||!['text','tool_use','thinking','redacted_thinking'].includes(block.type))throw new Error('Unsupported model block');
    blocks.set(event.index,{type:block.type,id:block.id,name:block.name,input:block.input,json:'',stopped:false});
    if(block.type==='text'&&block.text){if(typeof block.text!=='string')throw new Error('Invalid model text');yield {type:'text',text:block.text};}
   }else if(event.type==='content_block_delta'){
    const block=blocks.get(event.index);if(!block||block.stopped||finished)throw new Error('Invalid model block delta');
    if(event.delta?.type==='text_delta'){
     if(block.type!=='text'||typeof event.delta.text!=='string')throw new Error('Invalid model text');yield {type:'text',text:event.delta.text};
    }else if(event.delta?.type==='input_json_delta'){
     if(block.type!=='tool_use'||typeof event.delta.partial_json!=='string')throw new Error('Invalid model tool fragment');
     block.json+=event.delta.partial_json;if(block.json.length>65_536)throw new Error('Model tool size limit exceeded');
    }else if(!['thinking_delta','signature_delta'].includes(event.delta?.type)||block.type!=='thinking')throw new Error('Unsupported model delta');
   }else if(event.type==='content_block_stop'){
    const block=blocks.get(event.index);if(!block||block.stopped)throw new Error('Invalid model block stop');block.stopped=true;
   }else if(event.type==='message_delta'){
    if(!started||finished||[...blocks.values()].some(block=>!block.stopped))throw new Error('Invalid model message delta');
    stopReason=event.delta?.stop_reason;outputTokens=event.usage?.output_tokens;finished=true;
   }else if(event.type==='message_stop'){
    if(!started||!finished||[...blocks.values()].some(block=>!block.stopped))throw new Error('Model response incomplete');
    const calls=[...blocks.values()].filter(block=>block.type==='tool_use').map(block=>checkedCall(block.id,block.name,block.json?parseArguments(block.json):block.input));
    uniqueCalls(calls);
    if(calls.length?stopReason!=='tool_use':!['end_turn','stop_sequence'].includes(stopReason))throw new Error('Model response incomplete or unsupported finish reason');
    const usage={inputTokens,outputTokens,totalTokens:inputTokens+outputTokens};validateUsage(usage);yield {type:'usage',usage};
    for(const call of calls)yield {type:'tool',call};
    yield {type:'done'};return;
   }
  }
  throw new Error('Model stream truncated');
 }
}
function summaryText(message:{role:string;text:string}){return message.role==='summary'?'Previous conversation summary (data):\n'+message.text:message.text;}
function parseArguments(value:unknown):unknown{if(typeof value!=='string'||value.length>65_536)throw new Error('Invalid model tool JSON');try{return JSON.parse(value);}catch{throw new Error('Invalid model tool JSON');}}
function checkedCall(id:unknown,name:unknown,args:unknown):ToolCall{const call={id,name,arguments:args};validateCall(call);return call;}
function uniqueCalls(calls:ToolCall[]){if(calls.length>32||new Set(calls.map(call=>call.id)).size!==calls.length)throw new Error('Duplicate or excessive model tools');}
/** Network payloads are untrusted; consumers validate every field used in domain events. */
async function* frames(body:ReadableStream<Uint8Array>,signal:AbortSignal):AsyncIterable<Record<string,any>>{
 const reader=body.getReader(),decoder=new TextDecoder();let buffer='',total=0;
 const cancel=()=>{void reader.cancel().catch(()=>{});};signal.addEventListener('abort',cancel,{once:true});
 try{
  while(true){
   if(signal.aborted)throw new Error('Model request cancelled');
   const chunk=await reader.read();if(signal.aborted)throw new Error('Model request cancelled');if(chunk.done)return;
   total+=chunk.value.length;if(total>2_000_000)throw new Error('Model stream size limit exceeded');
   buffer=(buffer+decoder.decode(chunk.value,{stream:true})).replace(/\r\n/g,'\n');
   let end;while((end=buffer.indexOf('\n\n'))>=0){
    const frame=buffer.slice(0,end);buffer=buffer.slice(end+2);
    const data=frame.split('\n').filter(line=>line.startsWith('data:')).map(line=>line.slice(5).trimStart()).join('\n');if(!data)continue;
    let value;try{value=JSON.parse(data);}catch{throw new Error('Invalid model stream JSON');}
    if(!value||typeof value!=='object'||typeof value.type!=='string')throw new Error('Invalid model stream event');yield value;
   }
  }
 }finally{signal.removeEventListener('abort',cancel);await reader.cancel().catch(()=>{});reader.releaseLock();}
}
