import {resolve} from 'node:path';import type {Plugin} from 'vite';import {CloudServer} from './server/cloud';
import {serverConfig} from './server/config';
let cloud:CloudServer|undefined;
export function cloudUser(req:import('node:http').IncomingMessage){return cloud?.account(req);}
export function cloudPlugin():Plugin{return {name:'mydoc-cloud',configureServer(server){if(process.env.MYDOC_TEST_STORAGE==='1')return;const config=serverConfig();cloud=new CloudServer(config.cloudDbPath,{defaultAdmin:config.admin});server.middlewares.use('/api/cloud',cloud.handler);server.middlewares.use('/api/hosting',(_req,res)=>{res.setHeader('Content-Type','application/json');res.end(JSON.stringify({hosted:false}));});server.httpServer?.once('close',()=>{cloud?.close();cloud=undefined;});}};}
