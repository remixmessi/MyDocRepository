import {SQLiteSyncQueue} from './sync-queue';
import {randomBytes,randomUUID,scryptSync,timingSafeEqual,createHash} from 'node:crypto';
import type {DatabaseBackups} from './backups';
import type {IncomingMessage,ServerResponse} from 'node:http';
import type {SQLiteProjectDatabase} from './database';
const localAuthAttempts=new WeakMap<SQLiteProjectDatabase,Map<string,{count:number;until:number}>>();
function key(value:string|null){if(!value||value.length>1000||value.includes('\0'))throw new Error('存储标识无效');return value;}
async function body(req:IncomingMessage,max:number){const chunks:Buffer[]=[];let size=0;for await(const part of req){size+=part.length;if(size>max)throw new Error('存储请求超过大小限制');chunks.push(Buffer.from(part));}return Buffer.concat(chunks);}
export function storageHandler(database:SQLiteProjectDatabase,testMode=false,backups?:DatabaseBackups,scopeOverride?:string){return async(req:IncomingMessage,res:ServerResponse)=>{try{
 const requestGeneration=database.generation;const checkGeneration=()=>{if(database.restoring)throw new Error('数据库正在恢复，请稍后重试');if(database.generation!==requestGeneration||req.headers['x-mydoc-generation']&&req.headers['x-mydoc-generation']!==database.generation)throw new Error('数据库已恢复或服务已重启，请重新载入页面；尚未保存的内容请先复制。');};const mutate=<T>(action:()=>T)=>database.db.transaction(()=>{checkGeneration();return action();}).immediate();const readBody=async(max:number)=>{const result=await body(req,max);checkGeneration();return result;};
 const origin=req.headers.origin;if(origin&&new URL(origin).host!==req.headers.host){res.writeHead(403);res.end('禁止跨站存储访问');return;}
 const scope=scopeOverride??(testMode?String(req.headers['x-mydoc-test-scope']??'test'): 'main');key(scope);const url=new URL(req.url??'/','http://local'),route=url.pathname;res.setHeader('Cache-Control','no-store');const json=(value:unknown,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(value));};
 if(scopeOverride?.startsWith('user:')&&route!=='/status'&&req.headers['x-mydoc-account']!==scopeOverride.slice(5)){json({error:'账户已变更，请重新打开页面，未保存内容请先复制。'},409);return;}
 if(route!=='/status'&&!route.startsWith('/backups'))checkGeneration();
 if(route==='/local-account'&&req.method==='POST'){
  const input=JSON.parse((await readBody(4096)).toString());
  if(input.action==='login'||input.action==='register'){
   let attempts=localAuthAttempts.get(database);if(!attempts){attempts=new Map();localAuthAttempts.set(database,attempts);}
   const ip=req.socket.remoteAddress??'local',now=Date.now();for(const [key,value] of attempts)if(value.until<now)attempts.delete(key);
   const attempt=attempts.get(ip)??{count:0,until:now+60000};attempts.set(ip,attempt);if(++attempt.count>15){json({error:'尝试过于频繁，请稍后重试'},429);return;}
  }
  const digest=(value:string)=>createHash('sha256').update(value).digest('hex');
  if(['session','logout'].includes(input.action)){
   const token=typeof input.token==='string'?input.token:'';
   if(input.action==='logout'){database.db.prepare('DELETE FROM local_account_sessions WHERE token=?').run(digest(token));json({user:null});return;}
   const user=database.db.prepare('SELECT u.id,u.email,u.name FROM local_accounts u JOIN local_account_sessions s ON u.id=s.user WHERE s.token=? AND s.expires>?').get(digest(token),Date.now());json({user:user??null});return;
  }
  const email=typeof input.email==='string'?input.email.trim().toLowerCase():'',password=typeof input.password==='string'?input.password:'';
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||email.length>254||password.length<10||password.length>256){json({error:'请输入有效邮箱和 10 至 256 位密码'},400);return;}
  let row=database.db.prepare('SELECT * FROM local_accounts WHERE email=?').get(email) as {id:string;email:string;name:string;salt:string;password:string}|undefined;
  if(input.action==='register'){
   if(row){json({error:'本地账号已存在，请登录'},409);return;}
   const name=typeof input.name==='string'?input.name.trim():'';if(!name||name.length>160){json({error:'请输入显示名称（最多 160 字）'},400);return;}
   const salt=randomBytes(16).toString('hex');row={id:randomUUID(),email,name,salt,password:scryptSync(password,salt,64).toString('hex')};
   mutate(()=>database.db.prepare('INSERT INTO local_accounts VALUES (?,?,?,?,?)').run(row!.id,email,name,salt,row!.password));
  }else if(input.action!=='login'){json({error:'无效的账号操作'},400);return;}
  if(!row||!timingSafeEqual(scryptSync(password,row.salt,64),Buffer.from(row.password,'hex'))){json({error:'邮箱或密码不正确'},401);return;}
  const token=randomBytes(32).toString('hex');mutate(()=>database.db.prepare('INSERT INTO local_account_sessions VALUES (?,?,?)').run(digest(token),row!.id,Date.now()+7*86400000));
  json({user:{id:row.id,email:row.email,name:row.name},token});return;
 }
 if(route.startsWith('/backups')&&backups&&!testMode){if(route==='/backups'&&req.method==='GET'){json(backups.status());return;}if(route==='/backups/create'&&req.method==='POST'){checkGeneration();json(await backups.create());return;}if(route==='/backups/settings'&&req.method==='PUT'){json(backups.configure(JSON.parse((await readBody(4096)).toString())));return;}if(route==='/backups/preview'&&req.method==='GET'){json(await backups.preview(key(url.searchParams.get('id'))));return;}if(route==='/backups/restore'&&req.method==='POST'){checkGeneration();json(await backups.restore(key(url.searchParams.get('id'))));return;}}
 if(route==='/status'&&req.method==='GET'){json({accountId:scopeOverride?.startsWith('user:')?scopeOverride.slice(5):undefined,generation:database.generation,storage:'sqlite',initialized:true,schemaVersion:database.schemaVersion,empty:!database.db.prepare('SELECT 1 FROM documents WHERE scope=? LIMIT 1').get(scope),assetCount:(database.db.prepare('SELECT count(*) count FROM assets WHERE scope=?').get(scope) as {count:number}).count,preferences:Object.fromEntries((database.db.prepare('SELECT key,value FROM preferences WHERE scope=?').all(scope) as {key:string;value:string}[]).map(p=>[p.key,p.value]))});return;}
 if(route==='/sync-drafts'&&req.method==='GET'){const entries=database.db.prepare("SELECT id,server,workspace,doc,data,created,generation FROM drafts WHERE scope=? AND doc LIKE 'doc:%' ORDER BY created DESC LIMIT 1000").all(scope) as {id:string;server:string;workspace:string;doc:string;data:Uint8Array;created:number;generation:string}[];if(entries.reduce((n,row)=>n+row.data.length,0)>32*1024*1024)throw new Error('暂存草稿总量超过读取限制，请分批处理');json(entries.map(entry=>({...entry,data:Buffer.from(entry.data).toString('base64')})));return;}
 if(route==='/sync-queue'){
  const target={scope,server:key(url.searchParams.get('server')),workspace:key(url.searchParams.get('workspace')),doc:key(url.searchParams.get('id'))};
  const server=new URL(target.server);if(!['http:','https:'].includes(server.protocol)||server.origin!==target.server)throw new Error('同步服务器标识无效');
  if(!/^(doc:|workspace:)/.test(target.doc))throw new Error('同步文档标识无效');
  const queue=new SQLiteSyncQueue(database);
  if(req.method==='GET'){const rows=queue.list(target);json({generation:database.generation,rows:rows.map(row=>({...row,data:Buffer.from(row.data).toString('base64')})),lastSuccess:Number((database.db.prepare('SELECT value FROM preferences WHERE scope=? AND key=?').get(scope,'sync-last-success:'+JSON.stringify([target.server,target.workspace])) as {value:string}|undefined)?.value??0)});return;}
  if(req.method==='POST'){const bytes=await readBody(48*1024*1024);let update:Uint8Array=bytes,snapshot:Uint8Array|undefined;if(req.headers['content-type']==='application/json'){const input=JSON.parse(bytes.toString());const decode=(value:unknown)=>{if(typeof value!=='string'||!value||!/^[A-Za-z0-9+/]*={0,2}$/.test(value))throw new Error('队列数据无效');const result=Buffer.from(value,'base64');if(result.length>16*1024*1024)throw new Error('队列更新超过大小限制');return result;};update=decode(input.update);if(input.snapshot!==undefined)snapshot=decode(input.snapshot);}else if(bytes.length>16*1024*1024)throw new Error('队列更新超过大小限制');json({seq:mutate(()=>queue.enqueue(target,update,snapshot))});return;}
  if(req.method==='DELETE'||req.method==='PATCH'){const input=JSON.parse((await readBody(8192)).toString());if(!Number.isSafeInteger(input.seq)||input.seq<1)throw new Error('队列序号无效');mutate(()=>{if(req.method==='DELETE')queue.acknowledge(target,input.seq);else{if(typeof input.error!=='string')throw new Error('重试错误信息无效');queue.fail(target,input.seq,input.error);}});json({ok:true});return;}
 }
 if(route==='/document'){const id=key(url.searchParams.get('id'));if(req.method==='GET'){const bytes=database.load(scope,id);if(!bytes){res.writeHead(204);res.end();return;}res.writeHead(200,{'Content-Type':'application/octet-stream'});res.end(bytes);return;}if(req.method==='POST'){const bytes=await readBody(64*1024*1024);mutate(()=>database.append(scope,id,bytes));json({ok:true});return;}if(req.method==='DELETE'){mutate(()=>database.db.prepare('DELETE FROM documents WHERE scope=? AND id=?').run(scope,id));json({ok:true});return;}}
 if(route==='/asset'){const workspace=key(url.searchParams.get('workspace')),id=key(url.searchParams.get('id'));if(req.method==='GET'){const row=database.db.prepare('SELECT name,mime,data FROM assets WHERE scope=? AND workspace=? AND id=?').get(scope,workspace,id) as {name:string;mime:string;data:Buffer}|undefined;if(!row){json({error:'文件不存在'},404);return;}json({id,name:row.name,mime:row.mime,size:row.data.length,data:Buffer.from(row.data).toString('base64')});return;}if(req.method==='PUT'){const data=await readBody(20*1024*1024);const name=key(url.searchParams.get('name')),mime=key(url.searchParams.get('mime'));mutate(()=>database.db.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?) ON CONFLICT(scope,workspace,id) DO UPDATE SET name=excluded.name,mime=excluded.mime,data=excluded.data').run(scope,workspace,id,name,mime,data));json({ok:true});return;}if(req.method==='DELETE'){mutate(()=>database.db.prepare('DELETE FROM assets WHERE scope=? AND workspace=? AND id=?').run(scope,workspace,id));json({ok:true});return;}}
 if(route==='/assets'&&req.method==='DELETE'){mutate(()=>database.db.prepare('DELETE FROM assets WHERE scope=? AND workspace=?').run(scope,key(url.searchParams.get('workspace'))));json({ok:true});return;}
 if(route==='/preference'&&req.method==='PUT'){const input=(await readBody(4*1024*1024)).toString('utf8');const value=mutate(()=>database.preference(scope,key(url.searchParams.get('key')),input));json({ok:true,value});return;}
 if(route==='/test-clear-assets'&&testMode&&req.method==='DELETE'){database.db.prepare('DELETE FROM assets WHERE scope=?').run(scope);json({ok:true});return;}
 if(route==='/migrate'&&req.method==='POST'){const input=JSON.parse((await readBody(128*1024*1024)).toString()) as {documents:[string,string][];assets:{workspace:string;id:string;name:string;mime:string;data:string}[];preferences:[string,string][]};if(!Array.isArray(input.documents)||!Array.isArray(input.assets)||!Array.isArray(input.preferences)||input.documents.length>10000||input.assets.length>10000)throw new Error('迁移数据无效');const migrated=database.db.transaction(()=>{checkGeneration();if(database.db.prepare('SELECT 1 FROM documents WHERE scope=? LIMIT 1').get(scope))return false;for(const[id,data]of input.documents)database.append(scope,key(id),Buffer.from(data,'base64'));for(const a of input.assets){const bytes=Buffer.from(a.data,'base64');if(bytes.length>20*1024*1024)throw new Error('附件超过 20 MB');database.db.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?)').run(scope,key(a.workspace),key(a.id),key(a.name),key(a.mime),bytes);}for(const[k,v]of input.preferences)database.preference(scope,key(k),v);return true;}).immediate();json({migrated});return;}
 json({error:'存储接口不存在'},404);
 }catch(error){res.writeHead(500,{'Content-Type':'application/json'});res.end(JSON.stringify({error:error instanceof Error?error.message:'SQLite 写入失败'}));}};}
