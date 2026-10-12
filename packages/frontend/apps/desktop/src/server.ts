import {DatabaseBackups} from '../../../../common/nbstore/server/backups';
import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import {extname} from 'node:path';
import {randomBytes} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {storageHandler} from '../../../../common/nbstore/server/http';
import {clipperHandler} from '../../web/clipper-server';
import {staticFile} from './policy';
const types:Record<string,string>={'.html':'text/html; charset=utf-8','.js':'text/javascript','.mjs':'text/javascript','.css':'text/css','.json':'application/json','.wasm':'application/wasm','.svg':'image/svg+xml','.png':'image/png','.ico':'image/x-icon'};
export async function startDesktopServer(root:string,database:SQLiteProjectDatabase){
 const backups=new DatabaseBackups(database,database.connect);backups.start();
 const token=randomBytes(32).toString('hex'),cookieToken=randomBytes(32).toString('hex');
 const server=createServer(async(req,res)=>{
  try{
   const url=new URL(req.url??'/','http://local');
   if(req.headers.host!==origin.slice(7)){res.writeHead(403);res.end();return;}
   if(url.pathname.startsWith('/api/storage/')){if(req.headers.authorization!=='Bearer '+token){res.writeHead(403);res.end();return;}req.url=url.pathname.slice('/api/storage'.length)+url.search;await storageHandler(database,false,backups)(req,res);return;}
   if(url.pathname==='/'&&url.searchParams.get('desktopKey')===cookieToken){res.writeHead(303,{'Set-Cookie':`mydoc-desktop=${cookieToken}; HttpOnly; SameSite=Strict; Path=/`,'Location':'/'});res.end();return;}
   if(!req.headers.cookie?.split(';').some(c=>c.trim()==='mydoc-desktop='+cookieToken)){res.writeHead(403);res.end();return;}
   if(url.pathname==='/api/clipper'){await clipperHandler(req,res);return;}
   if(req.method!=='GET'&&req.method!=='HEAD'){res.writeHead(405);res.end();return;}
   const file=staticFile(root,req.url??'/');if(!file){res.writeHead(403);res.end();return;}
   const bytes=await readFile(file);res.setHeader('Content-Type',types[extname(file)]??'application/octet-stream');res.setHeader('Cache-Control','no-store');res.setHeader('X-Content-Type-Options','nosniff');
   res.setHeader('Content-Security-Policy',"default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' blob: data: https: http:; media-src 'self' blob:; worker-src 'self' blob:; connect-src 'self' blob: https: http:; font-src 'self' data: blob:; frame-src https: http:; object-src 'none'; base-uri 'none'");
   res.end(req.method==='HEAD'?undefined:bytes);
  }catch{if(!res.headersSent){res.writeHead(404);res.end();}}
 });
 let origin='';await new Promise<void>((resolve,reject)=>{server.once('error',reject);server.listen(0,'127.0.0.1',()=>resolve());});origin='http://127.0.0.1:'+(server.address() as {port:number}).port;
 return {backups,origin,url:origin+'/?desktopKey='+cookieToken,token,close:()=>new Promise<void>(resolve=>{backups.stop();server.close(()=>resolve());server.closeIdleConnections();})};
}
