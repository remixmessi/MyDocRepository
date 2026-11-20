import {resolve,join} from 'node:path';import type {Plugin} from 'vite';import {CloudServer} from './server/cloud';
import {mkdtempSync,rmSync} from 'node:fs';import {tmpdir} from 'node:os';import {build} from 'esbuild';
import type {AgentService} from './server/agent/service';
import {createRequire} from 'node:module';
import {serverConfig} from './server/config';
let cloud:CloudServer|undefined;
export function cloudUser(req:import('node:http').IncomingMessage){return cloud?.account(req);}
const disabledCloud=(req:import('node:http').IncomingMessage,res:import('node:http').ServerResponse)=>{res.setHeader('Content-Type','application/json');if(req.method==='GET'&&req.url?.split('?')[0]==='/session'){res.end(JSON.stringify({user:null,mail:{configured:false}}));return;}res.statusCode=404;res.end(JSON.stringify({error:'测试存储模式未启用团队服务器'}));};
export function cloudPlugin():Plugin{return {name:'mydoc-cloud',async configureServer(server){
 if(process.env.MYDOC_TEST_STORAGE==='1'){server.middlewares.use('/api/cloud',disabledCloud);server.middlewares.use('/api/agent',(_req,res)=>{res.statusCode=503;res.setHeader('Content-Type','application/json');res.end(JSON.stringify({error:'测试存储模式未启用 Agent 服务'}));});return;}
 const config=serverConfig(),owner=new CloudServer(config.cloudDbPath,{defaultAdmin:config.admin});cloud=owner;
 const runtime=await prepareAgentDevelopment();let agents:AgentService;
 try{agents=new runtime.AgentService(owner.database,runtime.worker,req=>{const user=owner.account(req);return user?{id:user.id,scope:'user:'+user.id}:undefined;});}
 catch(error){owner.close();runtime.cleanup();cloud=undefined;throw error;}
 server.middlewares.use('/api/cloud',owner.handler);server.middlewares.use('/api/agent',agents.handler);
 server.middlewares.use('/api/hosting',(_req,res)=>{res.setHeader('Content-Type','application/json');res.end(JSON.stringify({hosted:false}));});
 server.httpServer?.once('close',()=>{void agents.close().finally(()=>{owner.close();runtime.cleanup();if(cloud===owner)cloud=undefined;});});
}};}

export async function prepareAgentDevelopment(){
 const scratch=mkdtempSync(join(tmpdir(),'mydoc-agent-dev-')),worker=join(scratch,'agent-worker.cjs'),serviceEntry=join(scratch,'agent-service.cjs');
 try{await build({entryPoints:[resolve(__dirname,'../../../common/agent-host/src/worker-child.ts')],outfile:worker,bundle:true,platform:'node',format:'cjs'});await build({entryPoints:[resolve(__dirname,'server/agent/dev-entry.ts')],outfile:serviceEntry,bundle:true,platform:'node',format:'cjs',external:['better-sqlite3']});return {...createRequire(import.meta.url)(serviceEntry) as typeof import('./server/agent/dev-entry'),worker,cleanup:()=>rmSync(scratch,{recursive:true,force:true})};}catch(error){rmSync(scratch,{recursive:true,force:true});throw error;}
}
