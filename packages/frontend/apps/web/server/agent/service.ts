import {randomUUID} from 'node:crypto';
import type {IncomingMessage,ServerResponse} from 'node:http';
import type {SQLiteProjectDatabase} from '../../../../../common/nbstore/server/database';
import {SQLiteAgentStore} from '../../../../../common/agent-host/src/sqlite-store';
import {DurableAgentHost} from '../../../../../common/agent-host/src/durable-host';
import {ChildAgentRuntime} from '../../../../../common/agent-host/src/child-runtime';
import {CompatibleProvider} from '../../../../../common/agent-host/src/compatible-provider';
import {EnvSecrets,ModelConfigStore,ResilientModel,MODEL_DISCLOSURE,classifyModelError} from '../../../../../common/agent-host/src/model-config';
import type {SecretPort} from '../../../../../common/agent-host/src/model-config';
import type {ModelGateway} from '../../../../../common/agent-contracts/src/index';
export interface AgentPrincipal {scope:string;id:string}
/** userConfig: single-user desktop may save its own model; shared servers use administrator env only. */
export interface AgentServiceOptions {secrets?:SecretPort;userConfig?:boolean;transport?:typeof fetch}
/** Personal prompt-only API. Team documents/tools remain disabled until domain authorization is wired. */
export class AgentService {
 private hosts=new Map<string,{store:SQLiteAgentStore;host:DurableAgentHost}>();
 private jobs=new Map<string,Promise<void>>();private closing=false;
 private envProvider?:ModelGateway;private readonly config?:ModelConfigStore;private readonly secrets:SecretPort;private readonly transport?:typeof fetch;
 constructor(private readonly database:SQLiteProjectDatabase,private readonly worker:string,private readonly authenticate:(req:IncomingMessage)=>AgentPrincipal|undefined,env:NodeJS.ProcessEnv=process.env,options:AgentServiceOptions={}){
  this.secrets=options.secrets??new EnvSecrets(env);this.transport=options.transport;
  const allowed=env.MYDOC_AGENT_ALLOWED_HOSTS?.split(',').map(item=>item.trim().toLowerCase()).filter(Boolean);
  if(env.MYDOC_AGENT_ENDPOINT&&env.MYDOC_AGENT_MODEL){
   if(allowed&&!allowed.includes(new URL(env.MYDOC_AGENT_ENDPOINT).hostname.toLowerCase()))throw new Error('MYDOC_AGENT_ENDPOINT is not in MYDOC_AGENT_ALLOWED_HOSTS');
   this.envProvider=new ResilientModel(new CompatibleProvider({endpoint:env.MYDOC_AGENT_ENDPOINT,model:env.MYDOC_AGENT_MODEL,apiKey:new EnvSecrets(env).get('env:MYDOC_AGENT_API_KEY'),maxOutputTokens:Number(env.MYDOC_AGENT_MAX_OUTPUT_TOKENS??2048),contextWindow:Number(env.MYDOC_AGENT_CONTEXT_WINDOW??32000),supportsTools:env.MYDOC_AGENT_TOOLS==='1',includeUsage:env.MYDOC_AGENT_USAGE==='1'},options.transport));
  }
  if(options.userConfig)this.config=new ModelConfigStore(database,this.secrets,allowed);
  // This service is the sole owner during application startup, before listening.
  const scopes=this.database.db.prepare('SELECT DISTINCT scope FROM agent_sessions').all() as {scope:string}[];
  for(const {scope} of scopes)new SQLiteAgentStore(database,scope).recoverActive();
 }
 private owner(scope:string){let item=this.hosts.get(scope);if(!item){const store=new SQLiteAgentStore(this.database,scope);item={store,host:new DurableAgentHost(store,new ChildAgentRuntime(this.worker))};this.hosts.set(scope,item);}return item;}
 handler=async(req:IncomingMessage,res:ServerResponse)=>{
  const json=(value:unknown,status=200)=>{res.writeHead(status,{'Content-Type':'application/json','Cache-Control':'no-store'});res.end(JSON.stringify(value));};
  try{
   if(this.closing){json({error:'Agent service closing'},503);return;}
   const origin=req.headers.origin;
   if((origin&&new URL(origin).host!==req.headers.host)||req.headers['sec-fetch-site']==='cross-site'){json({error:'Cross-site request rejected'},403);return;}
   const principal=this.authenticate(req);if(!principal){json({error:'请先登录'},401);return;}
   const expected=req.headers['x-mydoc-account'];if(expected&&expected!==principal.id){json({error:'账户已变更'},409);return;}
   const check=()=>{const current=this.authenticate(req);if(!current||current.id!==principal.id||current.scope!==principal.scope)throw new Error('Account session expired');};
   const url=new URL(req.url??'/','http://local'),route=url.pathname.replace(/^\/api\/agent/,'');
   const {store,host}=this.owner(principal.scope);
   if(route==='/status'&&req.method==='GET'){const model=this.config?.publicView();json({enabled:!!this.envProvider||!!model,managed:!!this.envProvider,userConfig:!!this.config,secureStorage:this.config?this.secrets.available:false,model,disclosure:MODEL_DISCLOSURE,scope:'personal',worker:true});return;}
   if(route==='/sessions'&&req.method==='GET'){json(store.listSessions('personal'));return;}
   if(route==='/events'&&req.method==='GET'){json(store.replay(url.searchParams.get('session')??'',Number(url.searchParams.get('after')??0)));return;}
   if(route==='/runs'&&req.method==='GET'){const run=store.load(url.searchParams.get('id')??'');json(run?{...run,unconfirmed:store.unconfirmed(run.runId).map(call=>({id:call.id,name:call.name}))}:{error:'Run unavailable'},run?200:404);return;}
   if(req.method!=='POST'){json({error:'接口不存在'},404);return;}
   const generation=this.database.generation;
   if(req.headers['x-mydoc-generation']!==generation){json({error:'请重新载入后重试'},409);return;}
   const chunks:Buffer[]=[];let size=0;for await(const chunk of req){size+=chunk.length;if(size>110000)throw new Error('Request too large');chunks.push(Buffer.from(chunk));}
   const input=JSON.parse(Buffer.concat(chunks).toString());check();if(generation!==this.database.generation||this.database.restoring)throw new Error('Database generation changed');
   if(route==='/sessions'){
    if(typeof input.title!=='string'||!input.title.trim()||input.title.length>500)throw new Error('Invalid title');
    const count=this.database.db.prepare('SELECT count(*) n FROM agent_sessions WHERE scope=?').get(principal.scope) as {n:number};if(count.n>=100)throw new Error('Session limit reached');
    const id=randomUUID();store.createSession(id,'personal',input.title);json({id},201);return;
   }
   if(route==='/model'){
    if(!this.config||this.envProvider){json({error:'模型由管理员配置'},403);return;}
    if(input.clear===true)this.config.clear();else this.config.save(input);
    json({model:this.config.publicView()??null});return;
   }
   if(route==='/sessions/delete'){if(typeof input.id!=='string')throw new Error('Invalid session');store.deleteSession(input.id);json({deleted:true});return;}
   if(route==='/cancel'){if(!store.load(input.id))throw new Error('Run unavailable');json({cancelled:host.cancel(input.id)});return;}
   if(route==='/runs'){
    let provider:ModelGateway|undefined;
    try{provider=this.envProvider??this.config?.provider(this.transport);}catch(error){json({error:classifyModelError(error).message,code:classifyModelError(error).code},503);return;}
    if(!provider){json({error:this.config?'请先配置模型':'管理员尚未配置模型'},503);return;}
    const session=input.session;if(typeof session!=='string')throw new Error('Invalid session');
    const pending=this.database.db.prepare("SELECT count(*) n FROM agent_runs WHERE scope=? AND status='queued'").get(principal.scope) as {n:number};if(pending.n>=100)throw new Error('Queue limit reached');
    const jobKey=JSON.stringify([principal.scope,session]);if(!this.jobs.has(jobKey)&&this.jobs.size>=4){json({error:'Agent busy'},429);return;}
    const run=host.enqueue(session,randomUUID(),input.requestKey,input.prompt);
    if(!this.jobs.has(jobKey)){
     let activeRun:string|undefined;
     const monitor=setInterval(()=>{try{check();}catch{if(activeRun)host.cancel(activeRun);}},1000);monitor.unref();
     const job=(async()=>{while(!this.closing){check();const executed=await host.runNext(session,lease=>{activeRun=lease.runId;return {provider:{async *stream(messages,tools,signal){check();for await(const event of provider.stream(messages,tools,signal)){check();yield event;}}},tools:[],gateway:{async execute(){throw new Error('Tools disabled');}},clock:{now:()=>Date.now()},limits:{timeoutMs:120000},context:{inputBudget:24000}};});activeRun=undefined;if(!executed)break;}})().catch(()=>{/* Durable host journals failures; queued work awaits an authenticated retry. */}).finally(()=>{clearInterval(monitor);this.jobs.delete(jobKey);});
     this.jobs.set(jobKey,job);
    }
    json(run,202);return;
   }
   json({error:'接口不存在'},404);
  }catch{if(!res.headersSent)json({error:'Agent 请求失败，请检查会话、参数或运行状态'},400);}
 };
 async close(){this.closing=true;await Promise.all([...this.hosts.values()].map(item=>item.host.close()));await Promise.all(this.jobs.values());}
}
