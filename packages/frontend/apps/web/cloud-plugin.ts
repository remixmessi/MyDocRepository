import {resolve} from 'node:path';import type {Plugin} from 'vite';import {CloudServer} from './server/cloud';
import {serverConfig} from './server/config';
let cloud:CloudServer|undefined;
export function cloudUser(req:import('node:http').IncomingMessage){return cloud?.account(req);}
// With test storage the cloud server is off; /api/cloud answers "signed out" (session) or 404 as JSON instead of
// falling through to the HTML page, so the app does not report a broken team server.
const disabledCloud=(req:import('node:http').IncomingMessage,res:import('node:http').ServerResponse)=>{res.setHeader('Content-Type','application/json');if(req.method==='GET'&&req.url?.split('?')[0]==='/session'){res.end(JSON.stringify({user:null,mail:{configured:false}}));return;}res.statusCode=404;res.end(JSON.stringify({error:'测试存储模式未启用团队服务器'}));};
export function cloudPlugin():Plugin{return {name:'mydoc-cloud',configureServer(server){if(process.env.MYDOC_TEST_STORAGE==='1'){server.middlewares.use('/api/cloud',disabledCloud);return;}const config=serverConfig();cloud=new CloudServer(config.cloudDbPath,{defaultAdmin:config.admin});server.middlewares.use('/api/cloud',cloud.handler);server.middlewares.use('/api/hosting',(_req,res)=>{res.setHeader('Content-Type','application/json');res.end(JSON.stringify({hosted:false}));});server.httpServer?.once('close',()=>{cloud?.close();cloud=undefined;});}};}
