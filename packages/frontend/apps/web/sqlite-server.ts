import Database from 'better-sqlite3';
import {mkdirSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import type {IncomingMessage,ServerResponse} from 'node:http';
import type {Plugin} from 'vite';
import * as Y from 'yjs';
export const projectDatabasePath=resolve(dirname(fileURLToPath(import.meta.url)),'../../../../data/mydoc.sqlite');
export class ProjectDatabase {
 readonly db:Database.Database;
 readonly initializedThisStart:boolean;
 readonly schemaVersion=1;
 constructor(readonly path:string){
  mkdirSync(dirname(path),{recursive:true});this.db=new Database(path);
  try {
   this.db.pragma('busy_timeout = 5000');
   // Recheck under the write lock: concurrent starts must not initialize twice.
   this.initializedThisStart=this.db.transaction(()=>{
    const version=this.db.pragma('user_version',{simple:true}) as number;
    if(version>this.schemaVersion)throw new Error('数据库版本高于当前程序，请升级程序后再启动');
    if(version===this.schemaVersion)return false;
    this.db.exec(`CREATE TABLE IF NOT EXISTS documents(scope TEXT NOT NULL,id TEXT NOT NULL,data BLOB NOT NULL,PRIMARY KEY(scope,id));CREATE TABLE IF NOT EXISTS assets(scope TEXT NOT NULL,workspace TEXT NOT NULL,id TEXT NOT NULL,name TEXT NOT NULL,mime TEXT NOT NULL,data BLOB NOT NULL,PRIMARY KEY(scope,workspace,id));CREATE TABLE IF NOT EXISTS preferences(scope TEXT NOT NULL,key TEXT NOT NULL,value TEXT NOT NULL,PRIMARY KEY(scope,key));PRAGMA user_version=1;`);
    return true;
   }).immediate();
   this.db.pragma('journal_mode = WAL');this.db.pragma('synchronous = FULL');
  }catch(error){this.db.close();throw error;}
 }

 load(scope:string,id:string){return (this.db.prepare('SELECT data FROM documents WHERE scope=? AND id=?').get(scope,id) as {data:Buffer}|undefined)?.data??null;}
 append(scope:string,id:string,update:Uint8Array){const check=new Y.Doc();try{Y.applyUpdate(check,update);}finally{check.destroy();}this.db.transaction(()=>{const old=this.load(scope,id);const next=old?Y.mergeUpdates([old,update]):update;this.db.prepare('INSERT INTO documents VALUES (?,?,?) ON CONFLICT(scope,id) DO UPDATE SET data=excluded.data').run(scope,id,Buffer.from(next));}).immediate();}
 preference(scope:string,key:string,value:string){return this.db.transaction(()=>{if(key==='mydoc:reading-positions:v1'){const row=this.db.prepare('SELECT value FROM preferences WHERE scope=? AND key=?').get(scope,key) as {value:string}|undefined;let old:Record<string,{updatedAt:number}>={};try{old=JSON.parse(row?.value??'{}');}catch{}const incoming=JSON.parse(value) as typeof old;for(const[k,v]of Object.entries(incoming))if(!old[k]||v.updatedAt>=old[k].updatedAt)old[k]=v;value=JSON.stringify(Object.fromEntries(Object.entries(old).sort((a,b)=>b[1].updatedAt-a[1].updatedAt).slice(0,1000)));}this.db.prepare('INSERT INTO preferences VALUES (?,?,?) ON CONFLICT(scope,key) DO UPDATE SET value=excluded.value').run(scope,key,value);return value;}).immediate();}
 close(){this.db.close();}
}
function key(value:string|null){if(!value||value.length>1000||value.includes('\0'))throw new Error('存储标识无效');return value;}
async function body(req:IncomingMessage,max:number){const chunks:Buffer[]=[];let size=0;for await(const part of req){size+=part.length;if(size>max)throw new Error('存储请求超过大小限制');chunks.push(Buffer.from(part));}return Buffer.concat(chunks);}
export function storageHandler(database:ProjectDatabase,testMode=false){return async(req:IncomingMessage,res:ServerResponse)=>{try{
 const origin=req.headers.origin;if(origin&&new URL(origin).host!==req.headers.host){res.writeHead(403);res.end('禁止跨站存储访问');return;}
 const scope=testMode?String(req.headers['x-mydoc-test-scope']??'test'): 'main';key(scope);const url=new URL(req.url??'/','http://local'),route=url.pathname;res.setHeader('Cache-Control','no-store');const json=(value:unknown,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(value));};
 if(route==='/status'&&req.method==='GET'){json({storage:'sqlite',initialized:true,schemaVersion:database.schemaVersion,empty:!database.db.prepare('SELECT 1 FROM documents WHERE scope=? LIMIT 1').get(scope),assetCount:(database.db.prepare('SELECT count(*) count FROM assets WHERE scope=?').get(scope) as {count:number}).count,preferences:Object.fromEntries((database.db.prepare('SELECT key,value FROM preferences WHERE scope=?').all(scope) as {key:string;value:string}[]).map(p=>[p.key,p.value]))});return;}
 if(route==='/document'){const id=key(url.searchParams.get('id'));if(req.method==='GET'){const bytes=database.load(scope,id);if(!bytes){res.writeHead(204);res.end();return;}res.writeHead(200,{'Content-Type':'application/octet-stream'});res.end(bytes);return;}if(req.method==='POST'){database.append(scope,id,await body(req,64*1024*1024));json({ok:true});return;}if(req.method==='DELETE'){database.db.prepare('DELETE FROM documents WHERE scope=? AND id=?').run(scope,id);json({ok:true});return;}}
 if(route==='/asset'){const workspace=key(url.searchParams.get('workspace')),id=key(url.searchParams.get('id'));if(req.method==='GET'){const row=database.db.prepare('SELECT name,mime,data FROM assets WHERE scope=? AND workspace=? AND id=?').get(scope,workspace,id) as {name:string;mime:string;data:Buffer}|undefined;if(!row){json({error:'文件不存在'},404);return;}json({id,name:row.name,mime:row.mime,size:row.data.length,data:row.data.toString('base64')});return;}if(req.method==='PUT'){const data=await body(req,20*1024*1024);const name=key(url.searchParams.get('name')),mime=key(url.searchParams.get('mime'));database.db.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?) ON CONFLICT(scope,workspace,id) DO UPDATE SET name=excluded.name,mime=excluded.mime,data=excluded.data').run(scope,workspace,id,name,mime,data);json({ok:true});return;}if(req.method==='DELETE'){database.db.prepare('DELETE FROM assets WHERE scope=? AND workspace=? AND id=?').run(scope,workspace,id);json({ok:true});return;}}
 if(route==='/assets'&&req.method==='DELETE'){database.db.prepare('DELETE FROM assets WHERE scope=? AND workspace=?').run(scope,key(url.searchParams.get('workspace')));json({ok:true});return;}
 if(route==='/preference'&&req.method==='PUT'){const value=database.preference(scope,key(url.searchParams.get('key')),(await body(req,4*1024*1024)).toString('utf8'));json({ok:true,value});return;}
 if(route==='/test-clear-assets'&&testMode&&req.method==='DELETE'){database.db.prepare('DELETE FROM assets WHERE scope=?').run(scope);json({ok:true});return;}
 if(route==='/migrate'&&req.method==='POST'){const input=JSON.parse((await body(req,128*1024*1024)).toString()) as {documents:[string,string][];assets:{workspace:string;id:string;name:string;mime:string;data:string}[];preferences:[string,string][]};if(!Array.isArray(input.documents)||!Array.isArray(input.assets)||!Array.isArray(input.preferences)||input.documents.length>10000||input.assets.length>10000)throw new Error('迁移数据无效');const migrated=database.db.transaction(()=>{if(database.db.prepare('SELECT 1 FROM documents WHERE scope=? LIMIT 1').get(scope))return false;for(const[id,data]of input.documents)database.append(scope,key(id),Buffer.from(data,'base64'));for(const a of input.assets){const bytes=Buffer.from(a.data,'base64');if(bytes.length>20*1024*1024)throw new Error('附件超过 20 MB');database.db.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?)').run(scope,key(a.workspace),key(a.id),key(a.name),key(a.mime),bytes);}for(const[k,v]of input.preferences)database.preference(scope,key(k),v);return true;}).immediate();json({migrated});return;}
 json({error:'存储接口不存在'},404);
 }catch(error){res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({error:error instanceof Error?error.message:'SQLite 写入失败'}));}};}
export function sqliteServer():Plugin {let database:ProjectDatabase|undefined;const mount=(server:{middlewares:{use:(path:string,handler:(req:IncomingMessage,res:ServerResponse)=>void)=>void};httpServer?:{once:(event:string,callback:()=>void)=>void}|null})=>{database=new ProjectDatabase(process.env.MYDOC_DB_PATH??projectDatabasePath);server.middlewares.use('/api/storage',(req,res)=>void storageHandler(database!,process.env.MYDOC_TEST_STORAGE==='1')(req,res));server.httpServer?.once('close',()=>database?.close());};return {name:'mydoc-project-sqlite',configureServer:mount,configurePreviewServer:mount};}
