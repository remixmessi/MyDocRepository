import {mkdirSync,readFileSync,rmSync,writeFileSync,renameSync,existsSync} from 'node:fs';
import {join} from 'node:path';
import {randomUUID} from 'node:crypto';
import type {Message,ModelGateway,ProviderEvent,ToolSpec} from '@mydoc/agent-contracts';
import type {SQLiteProjectDatabase} from '../../nbstore/server/database';
import {createModelProvider} from './model-provider';
import type {ModelApiType} from '@mydoc/agent-contracts';
/** Secrets never enter SQLite, backups, renderer responses, logs or the worker; only a reference does. */
export interface SecretPort {readonly available:boolean;get(ref:string):string|undefined;put(value:string):string;delete(ref:string):void}
/** Server deployments: references name environment variables an administrator provisioned. */
export class EnvSecrets implements SecretPort {
 readonly available=false;
 constructor(private readonly env:NodeJS.ProcessEnv=process.env){}
 get(ref:string){return ref.startsWith('env:')?this.env[ref.slice(4)]:undefined;}
 put():string{throw new Error('服务端模型密钥只能由管理员通过环境配置');}
 delete(){}
}
export interface SecretCipher {isEncryptionAvailable():boolean;encryptString(value:string):Buffer;decryptString(value:Buffer):string}
/** Desktop: OS-backed encryption (Electron safeStorage). Refuses plaintext fallback. */
export class EncryptedFileSecrets implements SecretPort {
 constructor(private readonly directory:string,private readonly cipher:SecretCipher){}
 get available(){try{return this.cipher.isEncryptionAvailable();}catch{return false;}}
 private file(ref:string){if(!/^safe:[0-9a-f-]{36}$/.test(ref))throw new Error('Invalid secret reference');return join(this.directory,ref.slice(5)+'.bin');}
 get(ref:string){if(!ref.startsWith('safe:')||!this.available)return undefined;const file=this.file(ref);if(!existsSync(file))return undefined;try{return this.cipher.decryptString(readFileSync(file));}catch{return undefined;}}
 put(value:string){
  if(!this.available)throw new Error('系统加密不可用，无法安全保存模型密钥');
  const ref='safe:'+randomUUID(),file=this.file(ref);mkdirSync(this.directory,{recursive:true,mode:0o700});
  writeFileSync(file+'.tmp',this.cipher.encryptString(value),{mode:0o600});renameSync(file+'.tmp',file);return ref;
 }
 delete(ref:string){if(ref.startsWith('safe:'))rmSync(this.file(ref),{force:true});}
}
export interface ModelSettings {apiType?:ModelApiType;endpoint:string;model:string;maxOutputTokens:number;contextWindow:number;supportsTools:boolean;includeUsage:boolean;secretRef?:string}
export interface PublicModelSettings extends Omit<ModelSettings,'secretRef'> {hasKey:boolean;keyAvailable:boolean}
/** Shown on first configuration: exactly what leaves the device. */
export const MODEL_DISCLOSURE='启用后，提问内容、所选上下文、获准读取的资料片段和工具结果会发送到你配置的模型地址；密钥只保存在本机系统加密存储，不会写入数据库、备份或日志。';
export interface ModelConnection extends ModelSettings {id:string;name:string;models:string[];enabled:boolean}
export interface ModelConnections {version:1;connections:ModelConnection[];defaultModel?:{connectionId:string;model:string}}
const COLLECTION='agent.connections';
const KEY='agent.model',SCOPE='agent-config';
function host(endpoint:string){return new URL(endpoint).hostname.toLowerCase();}
export function validateSettings(value:unknown,allowedHosts?:readonly string[]):asserts value is ModelSettings {
 const s=value as ModelSettings;if(!s||typeof s!=='object')throw new Error('Invalid model settings');
 if(typeof s.endpoint!=='string'||s.endpoint.length>2000||typeof s.model!=='string'||typeof s.supportsTools!=='boolean'||typeof s.includeUsage!=='boolean')throw new Error('Invalid model settings');
 if(!Number.isSafeInteger(s.maxOutputTokens)||s.maxOutputTokens<1||s.maxOutputTokens>1_000_000||!Number.isSafeInteger(s.contextWindow)||s.contextWindow>10_000_000)throw new Error('Invalid model limits');
 if(s.secretRef!==undefined&&(typeof s.secretRef!=='string'||s.secretRef.length>300))throw new Error('Invalid secret reference');
 // Constructing validates protocol, loopback-only HTTP and the context window.
 createModelProvider({...s,apiKey:undefined});
 if(allowedHosts&&!allowedHosts.includes(host(s.endpoint)))throw new Error('模型地址不在允许列表中');
}
/** Host-owned configuration. The model and untrusted content can never choose an endpoint. */
export class ModelConfigStore {
 constructor(private readonly database:SQLiteProjectDatabase,private readonly secrets:SecretPort,private readonly allowedHosts?:readonly string[],private readonly scope=SCOPE){}
 private legacyLoad():ModelSettings|undefined{
  const row=this.database.db.prepare('SELECT value FROM preferences WHERE scope=? AND key=?').get(this.scope,KEY) as {value:string}|undefined;if(!row)return undefined;
  try{const value=JSON.parse(row.value);validateSettings(value,this.allowedHosts);return value;}catch{return undefined;}
 }
 connections():ModelConnections{
  const row=this.database.db.prepare('SELECT value FROM preferences WHERE scope=? AND key=?').get(this.scope,COLLECTION) as {value:string}|undefined;
  if(row){const value=JSON.parse(row.value) as ModelConnections;if(value.version!==1||!Array.isArray(value.connections))throw new Error('Unsupported model configuration');return value;}
  const old=this.legacyLoad();return {version:1,connections:old?[{...old,id:'legacy',name:'原有模型连接',models:[old.model],enabled:true}]:[],...(old?{defaultModel:{connectionId:'legacy',model:old.model}}:{})};
 }
 load():ModelSettings|undefined{
  const collection=this.connections(),selected=collection.defaultModel;
  const connection=collection.connections.find(item=>item.id===selected?.connectionId&&item.enabled&&item.models.includes(selected.model));
  if(!connection||!selected)return undefined;
  const {id,name,models,enabled,...settings}=connection;validateSettings(settings,this.allowedHosts);return {...settings,model:selected.model};
 }
 publicConnections(){const value=this.connections();return {...value,connections:value.connections.map(({secretRef,...item})=>({...item,hasKey:!!secretRef,keyAvailable:!secretRef||this.secrets.get(secretRef)!==undefined}))};}
 private writeConnections(value:ModelConnections){
  const selected=value.defaultModel,connection=value.connections.find(item=>item.id===selected?.connectionId);
  this.database.db.transaction(()=>{
   this.database.preference(this.scope,COLLECTION,JSON.stringify(value));
   if(connection&&selected){const {id,name,models,enabled,...settings}=connection;this.database.preference(this.scope,KEY,JSON.stringify({...settings,model:selected.model}));}
   else this.database.db.prepare('DELETE FROM preferences WHERE scope=? AND key=?').run(this.scope,KEY);
  }).immediate();
 }
 saveConnection(input:Omit<ModelConnection,'id'|'secretRef'>&{id?:string;apiKey?:string;keepKey?:boolean}){
  if(Object.hasOwn(input,'secretRef'))throw new Error('密钥引用只能由宿主生成');
  const value=this.connections(),old=value.connections.find(item=>item.id===input.id);
  if(input.id&&!old)throw new Error('连接不存在');
  if(typeof input.name!=='string'||!input.name.trim()||input.name.length>100||typeof input.enabled!=='boolean'||!Array.isArray(input.models)||!input.models.length||input.models.length>100||input.models.some(m=>typeof m!=='string'||!m.trim()||m.length>200)||new Set(input.models).size!==input.models.length||!input.models.includes(input.model))throw new Error('连接名称或模型列表无效');
  if(!old&&value.connections.length>=30)throw new Error('连接数量已达上限');
  const {id,name,models,enabled,apiKey,keepKey,...settings}=input;
  validateSettings(settings,this.allowedHosts);
  if(apiKey!==undefined&&(typeof apiKey!=='string'||apiKey.length>4096||/\s/.test(apiKey)))throw new Error('Invalid API key');
  if(keepKey&&old?.secretRef&&host(old.endpoint)!==host(settings.endpoint))throw new Error('更换模型地址时需要重新输入密钥');
  if(value.defaultModel&&value.defaultModel.connectionId===id&&(!enabled||!models.includes(value.defaultModel.model)))throw new Error('请先选择其他默认模型');
  const secretRef=apiKey?this.secrets.put(apiKey):keepKey?old?.secretRef:undefined;
  const connection:ModelConnection={...settings,id:old?.id??randomUUID(),name:name.trim(),models:[...models],enabled,...(secretRef?{secretRef}:{})};
  value.connections=old?value.connections.map(item=>item.id===old.id?connection:item):[...value.connections,connection];
  if(!value.defaultModel&&enabled)value.defaultModel={connectionId:connection.id,model:connection.model};
  try{this.writeConnections(value);}catch(error){if(secretRef&&secretRef!==old?.secretRef)this.secrets.delete(secretRef);throw error;}
  if(old?.secretRef&&old.secretRef!==secretRef)this.secrets.delete(old.secretRef);
  return connection.id;
 }
 setDefault(connectionId:string,model:string){const value=this.connections();if(!value.connections.some(item=>item.id===connectionId&&item.enabled&&item.models.includes(model)))throw new Error('默认模型不可用');value.defaultModel={connectionId,model};this.writeConnections(value);}
 deleteConnection(id:string){const value=this.connections(),old=value.connections.find(item=>item.id===id);if(!old)throw new Error('连接不存在');if(value.defaultModel?.connectionId===id)throw new Error('请先选择其他默认模型');value.connections=value.connections.filter(item=>item.id!==id);this.writeConnections(value);if(old.secretRef)this.secrets.delete(old.secretRef);}
 publicView():PublicModelSettings|undefined{const s=this.load();if(!s)return undefined;const {secretRef,...rest}=s;return {...rest,hasKey:!!secretRef,keyAvailable:!secretRef||this.secrets.get(secretRef)!==undefined};}
 save(input:Omit<ModelSettings,'secretRef'>&{apiKey?:string;keepKey?:boolean}){
  if(this.database.db.prepare('SELECT 1 FROM preferences WHERE scope=? AND key=?').get(this.scope,COLLECTION)){const value=this.connections(),old=value.connections.find(item=>item.id===value.defaultModel?.connectionId);if(!old)throw new Error('请通过连接管理添加模型');const {secretRef,...publicOld}=old;this.saveConnection({...publicOld,...input,models:[...new Set([...old.models,input.model])]});this.setDefault(old.id,input.model);return;}
  if(Object.hasOwn(input,'secretRef'))throw new Error('密钥引用只能由宿主生成');
  const {apiKey,keepKey,...settings}=input,old=this.load();
  if(apiKey!==undefined&&(typeof apiKey!=='string'||apiKey.length>4096||/\s/.test(apiKey)))throw new Error('Invalid API key');
  validateSettings(settings,this.allowedHosts);
  // A kept key is bound to its original host; changing the endpoint never forwards it elsewhere.
  if(keepKey&&old?.secretRef&&host(old.endpoint)!==host(settings.endpoint))throw new Error('更换模型地址时需要重新输入密钥');
  const secretRef=apiKey?this.secrets.put(apiKey):keepKey?old?.secretRef:undefined;
  this.database.preference(this.scope,KEY,JSON.stringify({...settings,...(secretRef?{secretRef}:{})}));
  if(old?.secretRef&&old.secretRef!==secretRef)this.secrets.delete(old.secretRef);
 }
 clear(){const refs=this.connections().connections.map(item=>item.secretRef);this.database.db.prepare('DELETE FROM preferences WHERE scope=? AND key IN (?,?)').run(this.scope,KEY,COLLECTION);for(const ref of refs)if(ref)this.secrets.delete(ref);}
 provider(transport?:typeof fetch):ModelGateway|undefined{
  const s=this.load();if(!s)return undefined;
  const apiKey=s.secretRef?this.secrets.get(s.secretRef):undefined;if(s.secretRef&&apiKey===undefined)throw new ModelError('auth','模型密钥不可用，请重新配置');
  return new ResilientModel(createModelProvider({...s,apiKey},transport));
 }
}
export type ModelErrorCode='auth'|'rate_limited'|'unavailable'|'network'|'timeout'|'invalid_response'|'cancelled';
export class ModelError extends Error {constructor(readonly code:ModelErrorCode,message:string,readonly retryable=false){super(message);}}
export function classifyModelError(error:unknown,signal?:AbortSignal):ModelError{
 if(error instanceof ModelError)return error;
 if(signal?.aborted)return new ModelError('cancelled','模型请求已取消');
 const message=error instanceof Error?error.message:'';const status=/Model HTTP status (\d+)/.exec(message)?.[1];
 if(status){const code=Number(status);
  if(code===401||code===403)return new ModelError('auth','模型认证失败，请检查密钥');
  if(code===429)return new ModelError('rate_limited','模型服务限流，请稍后重试',true);
  if(code>=500)return new ModelError('unavailable','模型服务暂不可用',true);
  return new ModelError('invalid_response','模型拒绝了请求（HTTP '+code+'）');
 }
 if(error instanceof TypeError||/fetch failed|ECONN|ENOTFOUND|EAI_AGAIN|socket/i.test(message))return new ModelError('network','无法连接模型服务',true);
 return new ModelError('invalid_response','模型响应无效');
}
export interface RetryPolicy {retries:number;baseDelayMs:number;firstEventTimeoutMs:number}
/** Bounded retry only before any event reached the caller; never switches provider or replays partial output. */
export class ResilientModel implements ModelGateway {
 constructor(private readonly inner:ModelGateway,private readonly policy:RetryPolicy={retries:2,baseDelayMs:500,firstEventTimeoutMs:60_000},private readonly sleep=(ms:number,signal:AbortSignal)=>new Promise<void>((resolve,reject)=>{const timer=setTimeout(resolve,ms);signal.addEventListener('abort',()=>{clearTimeout(timer);reject(new Error('cancelled'));},{once:true});})){}
 get capabilities(){return this.inner.capabilities;}
 async *stream(messages:readonly Message[],tools:readonly ToolSpec[],signal:AbortSignal):AsyncIterable<ProviderEvent>{
  for(let attempt=0;;attempt++){
   if(signal.aborted)throw new ModelError('cancelled','模型请求已取消');
   const controller=new AbortController(),abort=()=>controller.abort();signal.addEventListener('abort',abort,{once:true});
   // Race the first read as well as aborting, so an adapter that ignores its signal still times out.
   let timer:ReturnType<typeof setTimeout>|undefined;const timeout=new Promise<never>((_,reject)=>{timer=setTimeout(()=>{controller.abort();reject(new ModelError('timeout','模型响应超时',true));},this.policy.firstEventTimeoutMs);});
   const iterator=this.inner.stream(messages,tools,controller.signal)[Symbol.asyncIterator]();
   let yielded=false;
   try{
    while(true){
     const part=yielded?await iterator.next():await Promise.race([iterator.next(),timeout]);
     if(!yielded){yielded=true;clearTimeout(timer);}
     if(part.done)return;yield part.value;
    }
   }catch(error){
    const failure=classifyModelError(error,signal);
    if(yielded||!failure.retryable||attempt>=this.policy.retries||signal.aborted)throw failure;
   }finally{clearTimeout(timer);signal.removeEventListener('abort',abort);void Promise.resolve(iterator.return?.()).catch(()=>{});void timeout.catch(()=>{});}
   await this.sleep(this.policy.baseDelayMs*2**attempt,signal).catch(()=>{throw new ModelError('cancelled','模型请求已取消');});
  }
 }
}
