import Database from 'better-sqlite3';
import * as Y from 'yjs';
import {randomUUID,randomBytes,createHash,scryptSync,timingSafeEqual} from 'node:crypto';
import type {IncomingMessage,ServerResponse} from 'node:http';
import {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {MailService} from './mail';

const hash=(value:string)=>createHash('sha256').update(value).digest('hex');
type Account={id:string;email:string;name:string};type Member={user:string;role:'owner'|'admin'|'member'|'viewer'};
class RequestError extends Error {constructor(message:string,readonly status=400){super(message);}}
function text(value:unknown,max=160){if(typeof value!=='string'||!value.trim()||value.length>max)throw new RequestError('请输入有效内容');return value.trim();}
function position(value:unknown){if(!value||typeof value!=='object')return false;const p=value as {type?:{client:number;clock:number};item?:{client:number;clock:number};tname?:string;assoc?:number};const identifier=(v:typeof p.type)=>v==null||!!v&&Number.isSafeInteger(v.client)&&v.client>=0&&v.client<=0xffffffff&&Number.isSafeInteger(v.clock)&&v.clock>=0;return identifier(p.type)&&identifier(p.item)&&(p.tname==null||typeof p.tname==='string'&&p.tname.length<200)&&(p.assoc===undefined||Number.isSafeInteger(p.assoc));}
const AVATAR_TYPES=['image/png','image/jpeg','image/webp'],AVATAR_LIMIT=256*1024;
/** 按魔数判断图片类型：只认 PNG、JPEG、WebP，避免靠 Content-Type 就存任意字节。 */
function detectImage(data:Buffer){
 if(data.length>=8&&data.subarray(0,8).equals(Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a])))return 'image/png';
 if(data.length>=3&&data[0]===0xff&&data[1]===0xd8&&data[2]===0xff)return 'image/jpeg';
 if(data.length>=12&&data.subarray(0,4).toString('latin1')==='RIFF'&&data.subarray(8,12).toString('latin1')==='WEBP')return 'image/webp';
 return undefined;
}
async function body(req:IncomingMessage,max=8192){const chunks:Buffer[]=[];let size=0;for await(const chunk of req){size+=chunk.length;if(size>max)throw new RequestError('请求超过大小限制',413);chunks.push(Buffer.from(chunk));}return Buffer.concat(chunks);}
export class CloudServer {
 readonly database:SQLiteProjectDatabase;private streams=new Map<ServerResponse,{user:string;workspace:string;doc:string;expires:number;session:string;sessionId:string}>();private presence=new Map<string,{workspace:string;doc:string;client:number;clock:number;user:string;state:unknown;at:number}>();private attempts=new Map<string,{count:number;until:number}>();private timer:ReturnType<typeof setInterval>;
 readonly mail:MailService;
 constructor(path:string,options:{mail?:MailService}={}){this.mail=options.mail??new MailService();this.database=new SQLiteProjectDatabase(path,path=>new Database(path));this.database.db.exec(`CREATE TABLE IF NOT EXISTS cloud_deleted(workspace TEXT NOT NULL,id TEXT NOT NULL,PRIMARY KEY(workspace,id));CREATE TABLE IF NOT EXISTS cloud_users(id TEXT PRIMARY KEY,email TEXT UNIQUE NOT NULL,name TEXT NOT NULL,salt TEXT NOT NULL,password TEXT NOT NULL);CREATE TABLE IF NOT EXISTS cloud_sessions(token TEXT PRIMARY KEY,user TEXT NOT NULL,expires INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS cloud_workspaces(id TEXT PRIMARY KEY,name TEXT NOT NULL,created INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS cloud_members(workspace TEXT NOT NULL,user TEXT NOT NULL,role TEXT NOT NULL,PRIMARY KEY(workspace,user));CREATE TABLE IF NOT EXISTS cloud_invites(token TEXT PRIMARY KEY,workspace TEXT NOT NULL,email TEXT NOT NULL,role TEXT NOT NULL,expires INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS cloud_tokens(token TEXT PRIMARY KEY,user TEXT NOT NULL,kind TEXT NOT NULL,payload TEXT,created INTEGER NOT NULL,expires INTEGER NOT NULL);CREATE INDEX IF NOT EXISTS cloud_tokens_user ON cloud_tokens(user,kind);`);this.upgrade();this.timer=setInterval(()=>{for(const [res,state]of this.streams){if(state.expires<Date.now()||!this.database.db.prepare('SELECT 1 FROM cloud_sessions WHERE token=? AND expires>?').get(state.session,Date.now())||!this.member(state.workspace,state.user)){res.end();this.streams.delete(res);}else res.write(': heartbeat\n\n');}for(const[k,p]of this.presence)if(p.at<Date.now()-30000||!this.member(p.workspace,p.user)){this.broadcastPresence({...p,clock:p.clock+1,state:null});this.presence.delete(k);}for(const[k,v]of this.attempts)if(v.until<Date.now())this.attempts.delete(k);},15000);this.timer.unref();}
 // 协作表随账户设置扩列：逐列检查后 ALTER，已有库只补缺失列，不重建也不迁移数据。
 private upgrade(){const columns=(table:string)=>new Set((this.database.db.prepare(`PRAGMA table_info(${table})`).all() as {name:string}[]).map(column=>column.name));
  const add=(table:string,column:string,definition:string)=>{if(!columns(table).has(column))this.database.db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`);};
  this.database.db.transaction(()=>{
   add('cloud_users','avatar','BLOB');add('cloud_users','avatar_type','TEXT');add('cloud_users','verified','INTEGER NOT NULL DEFAULT 0');add('cloud_users','created','INTEGER NOT NULL DEFAULT 0');
   add('cloud_sessions','id','TEXT');add('cloud_sessions','created','INTEGER NOT NULL DEFAULT 0');add('cloud_sessions','seen','INTEGER NOT NULL DEFAULT 0');add('cloud_sessions','agent','TEXT');add('cloud_sessions','ip','TEXT');
   // 旧会话没有公开 ID，补一个随机值，避免把 token 哈希暴露给注销接口。
   for(const row of this.database.db.prepare('SELECT token FROM cloud_sessions WHERE id IS NULL').all() as {token:string}[])this.database.db.prepare('UPDATE cloud_sessions SET id=? WHERE token=?').run(randomUUID(),row.token);
  }).immediate();}
 private member(workspace:string,user:string){return this.database.db.prepare('SELECT user,role FROM cloud_members WHERE workspace=? AND user=?').get(workspace,user) as Member|undefined;}
 private token(req:IncomingMessage){return req.headers.cookie?.split(';').map(v=>v.trim()).find(v=>v.startsWith('mydoc_session='))?.slice(14);}
 account(req:IncomingMessage){const token=this.token(req);if(!token)return undefined;return this.database.db.prepare('SELECT u.id,u.email,u.name,u.verified,u.created AS createdAt,u.avatar_type AS avatarType,s.expires,s.id AS session FROM cloud_sessions s JOIN cloud_users u ON u.id=s.user WHERE s.token=? AND s.expires>?').get(hash(token),Date.now()) as (Account&{expires:number;verified:number;createdAt:number;avatarType:string|null;session:string})|undefined;}
 // 按 IP + 用途计数的频率限制，复用注册／登录原有的 1 分钟窗口。
 private limit(req:IncomingMessage,scope:string,max:number){const key=scope+':'+(req.socket.remoteAddress??'unknown'),limit=this.attempts.get(key)??{count:0,until:Date.now()+60000};if(limit.until<Date.now()){limit.count=0;limit.until=Date.now()+60000;}if(++limit.count>max)throw new RequestError('尝试过于频繁，请稍后重试',429);this.attempts.set(key,limit);}
 private cookie(res:ServerResponse,req:IncomingMessage,token:string){res.setHeader('Set-Cookie',`mydoc_session=${token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=604800${req.headers['x-forwarded-proto']==='https'?' ; Secure':''}`);}
 /** 建立会话并记录设备信息，供“登录设备”列表按会话注销。 */
 private startSession(user:string,req:IncomingMessage){const token=randomBytes(32).toString('hex'),now=Date.now();const agent=typeof req.headers['user-agent']==='string'?req.headers['user-agent'].slice(0,300):null;
  this.database.db.prepare('INSERT INTO cloud_sessions(token,user,expires,id,created,seen,agent,ip) VALUES (?,?,?,?,?,?,?,?)').run(hash(token),user,now+7*86400000,randomUUID(),now,now,agent,req.socket.remoteAddress??null);return token;}
 /** 一次性令牌：只存哈希，按用途覆盖同用户旧令牌，使用后立即删除。 */
 private issueToken(user:string,kind:'verify'|'reset'|'email',payload:string|null,ttl:number){const token=randomBytes(24).toString('hex'),now=Date.now();
  this.database.db.transaction(()=>{this.database.db.prepare('DELETE FROM cloud_tokens WHERE user=? AND kind=?').run(user,kind);this.database.db.prepare('INSERT INTO cloud_tokens(token,user,kind,payload,created,expires) VALUES (?,?,?,?,?,?)').run(hash(token),user,kind,payload,now,now+ttl);}).immediate();return token;}
 private consumeToken(token:string,kind:'verify'|'reset'|'email'){const row=this.database.db.prepare('SELECT user,payload,expires FROM cloud_tokens WHERE token=? AND kind=?').get(hash(token),kind) as {user:string;payload:string|null;expires:number}|undefined;
  if(!row){throw new RequestError('口令无效或已被使用');}
  this.database.db.prepare('DELETE FROM cloud_tokens WHERE token=?').run(hash(token));
  if(row.expires<Date.now())throw new RequestError('口令已过期，请重新获取');
  return row;}
 private setPassword(user:string,password:string){const salt=randomBytes(16).toString('hex');this.database.db.prepare('UPDATE cloud_users SET salt=?,password=? WHERE id=?').run(salt,scryptSync(password,salt,64).toString('hex'),user);}
 /** 口令投递：配置了 SMTP 就发邮件，未配置时只回传给已验明身份的当前用户，不假装已发送。 */
 private async deliver(to:string,subject:string,text:string){if(!this.mail.status.configured)return {sent:false,reason:this.mail.status.reason??'邮件服务未配置'};
  try{await this.mail.send({to,subject,text});return {sent:true};}catch(error){return {sent:false,reason:error instanceof Error?error.message:'邮件发送失败'};}}
 private require(req:IncomingMessage){const user=this.account(req);if(!user)throw new RequestError('请先登录',401);const expected=req.headers['x-mydoc-account']??new URL(req.url??'/','http://local').searchParams.get('account');if(expected&&expected!==user.id)throw new RequestError('账户已变更，请重新打开页面',409);return user;}
 private authorize(workspace:string,user:string,admin=false){const member=this.member(workspace,user);if(!member)throw new RequestError('你已不在此工作空间中',403);if(admin&&!['owner','admin'].includes(member.role))throw new RequestError('需要工作空间管理员权限',403);return member;}
 private writable(workspace:string,req:IncomingMessage){const user=this.require(req);const member=this.authorize(workspace,user.id);if(member.role==='viewer')throw new RequestError('查看者只能阅读工作空间',403);return member;}
 private push(workspace:string,doc:string,data:Uint8Array){for(const[res,s]of this.streams)if(s.workspace===workspace&&(s.doc===doc||s.doc==='*')){if(!res.write(`data: ${JSON.stringify({id:doc,data:Buffer.from(data).toString('base64')})}\n\n`)){res.end();this.streams.delete(res);}}}
 private broadcastPresence(p:{workspace:string;doc:string;client:number;clock:number;state:unknown}){for(const[res,s]of this.streams)if(s.workspace===p.workspace)res.write(`data: ${JSON.stringify({id:p.doc,presence:{client:p.client,clock:p.clock,state:p.state}})}\n\n`);}
 private disconnect(workspace:string,user?:string){for(const[res,s]of this.streams)if(s.workspace===workspace&&(!user||s.user===user)){res.end();this.streams.delete(res);}}
 /** 断开某账户的全部实时连接：注销、改密与重置密码后立即生效，不等心跳。 */
 private dropUserStreams(user:string,session?:string){for(const[res,s]of this.streams)if(s.user===user&&(!session||s.sessionId===session)){res.end();this.streams.delete(res);}}
 handler=async(req:IncomingMessage,res:ServerResponse)=>{const json=(value:unknown,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(value));};try{
 res.setHeader('Cache-Control','no-store');const url=new URL(req.url??'/','http://local'),route=url.pathname;const origin=req.headers.origin;if(origin&&new URL(origin).host!==req.headers.host)throw new RequestError('禁止跨站请求',403);if(req.headers['sec-fetch-site']==='cross-site')throw new RequestError('禁止跨站请求',403);
 // 会话信息附带邮箱验证状态、到期时间与邮件服务是否可用，界面据此提示而不是猜测。
 if(route==='/session'&&req.method==='GET'){const current=this.account(req);if(current)this.database.db.prepare('UPDATE cloud_sessions SET seen=? WHERE id=?').run(Date.now(),current.session);
  json({user:current?{id:current.id,email:current.email,name:current.name,verified:!!current.verified,createdAt:current.createdAt,hasAvatar:!!current.avatarType,expiresAt:current.expires,session:current.session}:null,mail:{configured:this.mail.status.configured,reason:this.mail.status.reason}});return;}
 if((route==='/register'||route==='/login')&&req.method==='POST'){this.limit(req,'auth',15);const input=JSON.parse((await body(req)).toString()),email=text(input.email,254).toLowerCase(),password=text(input.password,256);if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||password.length<10)throw new RequestError('请输入有效邮箱及至少 10 位密码');let user=this.database.db.prepare('SELECT * FROM cloud_users WHERE email=?').get(email) as (Account&{salt:string;password:string;verified:number})|undefined;if(route==='/register'){if(user)throw new RequestError('无法创建账户，请尝试登录');const id=randomUUID(),name=text(input.name),salt=randomBytes(16).toString('hex');this.database.db.prepare('INSERT INTO cloud_users(id,email,name,salt,password,verified,created) VALUES (?,?,?,?,?,0,?)').run(id,email,name,salt,scryptSync(password,salt,64).toString('hex'),Date.now());user={id,email,name,salt,password:'',verified:0};}else{const salt=user?.salt??'00000000000000000000000000000000';const actual=scryptSync(password,salt,64),expected=user?Buffer.from(user.password,'hex'):Buffer.alloc(64);if(!timingSafeEqual(actual,expected)||!user)throw new RequestError('邮箱或密码不正确',401);}this.cookie(res,req,this.startSession(user.id,req));json({user:{id:user.id,email:user.email,name:user.name,verified:!!user.verified}});return;}
 // 找回密码：未登录即可请求，响应不透露邮箱是否存在；未配置邮件服务时口令只写服务端日志，不回传给请求方。
 if(route==='/forgot'&&req.method==='POST'){this.limit(req,'forgot',5);const input=JSON.parse((await body(req)).toString()),email=text(input.email,254).toLowerCase();
  const target=this.database.db.prepare('SELECT id,email FROM cloud_users WHERE email=?').get(email) as {id:string;email:string}|undefined;
  let delivery:{sent:boolean;reason?:string}={sent:false,reason:this.mail.status.reason};
  if(target){const token=this.issueToken(target.id,'reset',null,3600000);
   delivery=await this.deliver(target.email,'MyDoc 密码重置口令','使用以下一次性口令重置密码，1 小时内有效，只能使用一次：\n\n'+token+'\n\n如果不是你本人申请，可忽略此邮件。');
   if(!delivery.sent)console.warn('[mydoc] 邮件服务不可用（'+(delivery.reason??'未知原因')+'），'+target.email+' 的密码重置口令：'+token);}
  json({ok:true,mailConfigured:this.mail.status.configured,...(this.mail.status.configured?{}:{notice:'邮件服务未配置（'+(this.mail.status.reason??'缺少 SMTP 设置')+'），重置口令已写入服务端日志，请联系运维获取。'})});return;}
 if(route==='/reset'&&req.method==='POST'){this.limit(req,'reset',10);const input=JSON.parse((await body(req)).toString()),token=text(input.token,128),password=text(input.password,256);
  if(password.length<10)throw new RequestError('密码至少 10 位');
  const row=this.consumeToken(token,'reset');
  // 重置后清空该账户全部会话：旧设备一律需要用新密码重新登录。
  this.database.db.transaction(()=>{this.setPassword(row.user,password);this.database.db.prepare('DELETE FROM cloud_sessions WHERE user=?').run(row.user);}).immediate();
  this.dropUserStreams(row.user);json({ok:true});return;}
 const user=this.require(req);
 if(route==='/logout'&&req.method==='POST'){const token=this.token(req);if(token)this.database.db.prepare('DELETE FROM cloud_sessions WHERE token=?').run(hash(token));this.disconnect('',user.id);this.dropUserStreams(user.id);res.setHeader('Set-Cookie','mydoc_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0');json({ok:true});return;}
 // 显示名称：头像单独走 /avatar，避免把二进制塞进 JSON 请求体。
 if(route==='/profile'&&req.method==='PATCH'){const input=JSON.parse((await body(req)).toString()),name=text(input.name);this.database.db.prepare('UPDATE cloud_users SET name=? WHERE id=?').run(name,user.id);json({ok:true,name});return;}
 // 头像限制 256 KiB、仅 PNG／JPEG／WebP，按魔数校验而不是只看 Content-Type。
 if(route==='/avatar'&&req.method==='PUT'){const type=String(req.headers['content-type']??'').split(';')[0].trim();
  if(!AVATAR_TYPES.includes(type))throw new RequestError('头像仅支持 PNG、JPEG 或 WebP');
  const data=await body(req,AVATAR_LIMIT);if(!data.length)throw new RequestError('头像内容为空');
  if(detectImage(data)!==type)throw new RequestError('头像内容与声明的图片格式不一致');
  this.database.db.prepare('UPDATE cloud_users SET avatar=?,avatar_type=? WHERE id=?').run(data,type,user.id);json({ok:true,type,bytes:data.length});return;}
 if(route==='/avatar'&&req.method==='DELETE'){this.database.db.prepare('UPDATE cloud_users SET avatar=NULL,avatar_type=NULL WHERE id=?').run(user.id);json({ok:true});return;}
 // 头像读取开放给已登录账户：成员列表本来就能看到姓名与邮箱，头像不额外提权。
 if(route==='/avatar'&&req.method==='GET'){const id=url.searchParams.get('user')??user.id;
  const row=this.database.db.prepare('SELECT avatar,avatar_type FROM cloud_users WHERE id=?').get(id) as {avatar:Buffer|null;avatar_type:string|null}|undefined;
  if(!row?.avatar||!row.avatar_type){res.writeHead(204);res.end();return;}
  res.writeHead(200,{'Content-Type':row.avatar_type,'Content-Length':String(row.avatar.length),'Cache-Control':'no-store'});res.end(row.avatar);return;}
 // 改密校验当前密码；默认让其他设备全部重新登录，当前会话换新 token 继续可用。
 if(route==='/password'&&req.method==='POST'){this.limit(req,'password',10);const input=JSON.parse((await body(req)).toString()),current=text(input.current,256),next=text(input.password,256);
  if(next.length<10)throw new RequestError('新密码至少 10 位');
  if(next===current)throw new RequestError('新密码与当前密码相同');
  const row=this.database.db.prepare('SELECT salt,password FROM cloud_users WHERE id=?').get(user.id) as {salt:string;password:string};
  if(!timingSafeEqual(scryptSync(current,row.salt,64),Buffer.from(row.password,'hex')))throw new RequestError('当前密码不正确',401);
  const keepOthers=input.signOutOthers===false;const token=randomBytes(32).toString('hex'),now=Date.now();
  this.database.db.transaction(()=>{this.setPassword(user.id,next);
   if(!keepOthers)this.database.db.prepare('DELETE FROM cloud_sessions WHERE user=? AND id<>?').run(user.id,user.session);
   // 当前会话也换新 token：旧 Cookie 失效，但本页面不被登出。
   this.database.db.prepare('UPDATE cloud_sessions SET token=?,seen=? WHERE id=?').run(hash(token),now,user.session);}).immediate();
  if(!keepOthers)for(const[stream,s]of this.streams)if(s.user===user.id&&s.sessionId!==user.session){stream.end();this.streams.delete(stream);}
  this.cookie(res,req,token);json({ok:true,signedOutOthers:!keepOthers});return;}
 // 登录设备列表：只暴露会话公开 ID，不返回 token 哈希。
 if(route==='/sessions'&&req.method==='GET'){const rows=this.database.db.prepare('SELECT id,created AS createdAt,seen AS lastSeenAt,expires AS expiresAt,agent,ip FROM cloud_sessions WHERE user=? AND expires>? ORDER BY seen DESC').all(user.id,Date.now()) as {id:string}[];
  json(rows.map(row=>({...row,current:row.id===user.session})));return;}
 if(route==='/sessions'&&req.method==='DELETE'){const input=JSON.parse((await body(req)).toString()) as {all?:boolean;session?:unknown};
  if(input.all===true){const removed=this.database.db.prepare('DELETE FROM cloud_sessions WHERE user=? AND id<>?').run(user.id,user.session) as {changes:number};
   for(const[stream,s]of this.streams)if(s.user===user.id&&s.sessionId!==user.session){stream.end();this.streams.delete(stream);}
   json({ok:true,removed:removed.changes});return;}
  const id=text(input.session,100);
  if(id===user.session)throw new RequestError('这是当前设备的会话，请使用“退出登录”');
  const removed=this.database.db.prepare('DELETE FROM cloud_sessions WHERE user=? AND id=?').run(user.id,id) as {changes:number};
  if(!removed.changes)throw new RequestError('会话不存在或已失效');
  for(const[stream,s]of this.streams)if(s.user===user.id&&s.sessionId===id){stream.end();this.streams.delete(stream);}
  json({ok:true,removed:removed.changes});return;}
 // 邮箱验证：发送／重发验证口令，1 小时有效，按 IP 限频。
 if(route==='/verify'&&req.method==='POST'){if(user.verified)throw new RequestError('邮箱已验证');this.limit(req,'verify',5);
  const token=this.issueToken(user.id,'verify',null,3600000);
  const delivery=await this.deliver(user.email,'MyDoc 邮箱验证口令','使用以下一次性口令完成邮箱验证，1 小时内有效：\n\n'+token+'\n');
  // 未配置邮件服务时把口令回传给本人：请求者已通过会话验明身份，收件地址就是其当前邮箱。
  json({ok:true,sent:delivery.sent,...(delivery.sent?{}:{token,notice:'邮件未发送（'+(delivery.reason??'邮件服务未配置')+'），以下口令仅本次显示。'})});return;}
 if(route==='/verify'&&req.method==='PATCH'){const input=JSON.parse((await body(req)).toString()),token=text(input.token,128);
  const row=this.consumeToken(token,'verify');
  if(row.user!==user.id)throw new RequestError('口令与当前账户不符');
  this.database.db.prepare('UPDATE cloud_users SET verified=1 WHERE id=?').run(user.id);json({ok:true});return;}
 // 更改邮箱：先向新地址发验证口令，确认后才写入；未接受的邀请按邮箱匹配，一并提示影响。
 if(route==='/email'&&req.method==='POST'){this.limit(req,'email',5);const input=JSON.parse((await body(req)).toString()),email=text(input.email,254).toLowerCase(),password=text(input.password,256);
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new RequestError('邮箱无效');
  if(email===user.email)throw new RequestError('新邮箱与当前邮箱相同');
  const row=this.database.db.prepare('SELECT salt,password FROM cloud_users WHERE id=?').get(user.id) as {salt:string;password:string};
  if(!timingSafeEqual(scryptSync(password,row.salt,64),Buffer.from(row.password,'hex')))throw new RequestError('当前密码不正确',401);
  if(this.database.db.prepare('SELECT 1 FROM cloud_users WHERE email=?').get(email))throw new RequestError('该邮箱已被其他账户使用');
  const pending=this.database.db.prepare('SELECT count(*) count FROM cloud_invites WHERE email=? AND expires>?').get(user.email,Date.now()) as {count:number};
  const token=this.issueToken(user.id,'email',email,3600000);
  const delivery=await this.deliver(email,'MyDoc 邮箱更改验证','使用以下一次性口令确认把账户邮箱改为 '+email+'，1 小时内有效：\n\n'+token+'\n');
  json({ok:true,sent:delivery.sent,email,pendingInvites:pending.count,...(delivery.sent?{}:{token,notice:'邮件未发送（'+(delivery.reason??'邮件服务未配置')+'），以下口令仅本次显示。'})});return;}
 if(route==='/email'&&req.method==='PATCH'){const input=JSON.parse((await body(req)).toString()),token=text(input.token,128);
  const row=this.consumeToken(token,'email');
  if(row.user!==user.id)throw new RequestError('口令与当前账户不符');
  const email=row.payload;if(!email)throw new RequestError('口令缺少目标邮箱，请重新申请');
  if(this.database.db.prepare('SELECT 1 FROM cloud_users WHERE email=? AND id<>?').get(email,user.id))throw new RequestError('该邮箱已被其他账户使用');
  // 发往旧地址的未接受邀请不迁移也不删除：邀请绑定邮箱本身，换邮箱后需重新邀请。
  const stale=this.database.db.prepare('SELECT count(*) count FROM cloud_invites WHERE email=? AND expires>?').get(user.email,Date.now()) as {count:number};
  this.database.db.prepare('UPDATE cloud_users SET email=?,verified=1 WHERE id=?').run(email,user.id);
  json({ok:true,email,staleInvites:stale.count});return;}
 if(route==='/workspaces'&&req.method==='GET'){json(this.database.db.prepare('SELECT w.id,w.name,w.created AS createdAt,m.role FROM cloud_workspaces w JOIN cloud_members m ON w.id=m.workspace WHERE m.user=? ORDER BY w.created').all(user.id));return;}
 if(route==='/workspaces'&&req.method==='POST'){const input=JSON.parse((await body(req)).toString()),id=randomUUID(),name=text(input.name),createdAt=Date.now();this.database.db.transaction(()=>{this.database.db.prepare('INSERT INTO cloud_workspaces VALUES (?,?,?)').run(id,name,createdAt);this.database.db.prepare('INSERT INTO cloud_members VALUES (?,?,?)').run(id,user.id,'owner');}).immediate();json({id,name,createdAt,role:'owner'},201);return;}
 if(route==='/accept'&&req.method==='POST'){const token=text(JSON.parse((await body(req)).toString()).token,128);const invite=this.database.db.prepare('SELECT * FROM cloud_invites WHERE token=? AND email=? AND expires>?').get(hash(token),user.email,Date.now()) as {workspace:string;role:string}|undefined;if(!invite)throw new RequestError('邀请无效、已使用、已过期或与当前邮箱不符');this.database.db.transaction(()=>{this.database.db.prepare('INSERT OR IGNORE INTO cloud_members VALUES (?,?,?)').run(invite.workspace,user.id,invite.role);this.database.db.prepare('DELETE FROM cloud_invites WHERE token=?').run(hash(token));}).immediate();json({ok:true});return;}
 const workspace=text(url.searchParams.get('workspace'),100);let member=this.authorize(workspace,user.id);
 if(route==='/access'&&req.method==='GET'){json({role:member.role});return;}
 // 空间概览：名称、所有者、当前角色、成员数量与存储占用，供设置页集中展示。
 if(route==='/overview'&&req.method==='GET'){const space=this.database.db.prepare('SELECT name,created FROM cloud_workspaces WHERE id=?').get(workspace) as {name:string;created:number};const owner=this.database.db.prepare("SELECT u.name,u.email FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=? AND m.role='owner'").get(workspace) as {name:string;email:string}|undefined;const counts=this.database.db.prepare("SELECT count(*) total,sum(role='owner') owners,sum(role='admin') admins,sum(role='member') members,sum(role='viewer') viewers FROM cloud_members WHERE workspace=?").get(workspace) as {total:number;owners:number;admins:number;members:number;viewers:number};const documents=this.database.db.prepare("SELECT count(*) count,coalesce(sum(length(data)),0) bytes FROM documents WHERE scope=? AND id LIKE 'doc:%'").get('cloud:'+workspace) as {count:number;bytes:number};const shared=this.database.db.prepare("SELECT coalesce(sum(length(data)),0) bytes FROM documents WHERE scope=? AND id LIKE 'workspace:%'").get('cloud:'+workspace) as {bytes:number};const assets=this.database.db.prepare('SELECT count(*) count,coalesce(sum(length(data)),0) bytes FROM assets WHERE scope=?').get('cloud:'+workspace) as {count:number;bytes:number};const invites=this.database.db.prepare('SELECT count(*) count FROM cloud_invites WHERE workspace=? AND expires>?').get(workspace,Date.now()) as {count:number};json({id:workspace,name:space.name,createdAt:space.created,role:member.role,owner:owner?{name:owner.name,email:owner.email}:null,members:counts,documents:documents.count,invites:invites.count,storage:{documents:documents.bytes+shared.bytes,assets:assets.bytes,assetCount:assets.count,total:documents.bytes+shared.bytes+assets.bytes}});return;}
 if(route==='/workspace'&&req.method==='PATCH'){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString());this.authorize(workspace,this.require(req).id,true);this.database.db.prepare('UPDATE cloud_workspaces SET name=? WHERE id=?').run(text(input.name),workspace);json({ok:true});return;}
 if(route==='/workspace'&&req.method==='DELETE'){if(member.role!=='owner')throw new RequestError('只有所有者可以删除工作空间',403);this.database.db.transaction(()=>{for(const table of ['cloud_members','cloud_invites','cloud_deleted'])this.database.db.prepare(`DELETE FROM ${table} WHERE workspace=?`).run(workspace);this.database.db.prepare('DELETE FROM cloud_workspaces WHERE id=?').run(workspace);this.database.db.prepare('DELETE FROM documents WHERE scope=?').run('cloud:'+workspace);this.database.db.prepare('DELETE FROM assets WHERE scope=?').run('cloud:'+workspace);}).immediate();this.disconnect(workspace);json({ok:true});return;}
 if(route==='/members'&&req.method==='GET'){json(this.database.db.prepare('SELECT u.id,u.name,u.email,m.role FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=?').all(workspace));return;}
 if(route==='/invite'&&req.method==='POST'){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString()),email=text(input.email,254).toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new RequestError('邮箱无效');const token=randomBytes(24).toString('hex');this.authorize(workspace,this.require(req).id,true);const role=input.role==='viewer'?'viewer':'member',expiresAt=Date.now()+7*86400000;
  // 重新生成即替换：同邮箱的旧口令一并删除，避免旧口令继续可用。
  this.database.db.transaction(()=>{this.database.db.prepare('DELETE FROM cloud_invites WHERE workspace=? AND email=?').run(workspace,email);this.database.db.prepare('INSERT INTO cloud_invites VALUES (?,?,?,?,?)').run(hash(token),workspace,email,role,expiresAt);}).immediate();json({token,role,expiresAt});return;}
 if(route==='/member'&&['PATCH','DELETE'].includes(req.method??'')){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString()),target=text(input.user,100),old=this.member(workspace,target);member=this.authorize(workspace,this.require(req).id,true);if(!old||old.role==='owner'||member.role!=='owner'&&old.role==='admin')throw new RequestError('不能修改此成员',403);if(req.method==='DELETE'){this.database.db.prepare('DELETE FROM cloud_members WHERE workspace=? AND user=?').run(workspace,target);this.disconnect(workspace,target);}else{if(member.role!=='owner'||!['admin','member','viewer'].includes(input.role))throw new RequestError('角色无效或权限不足',403);this.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run(input.role,workspace,target);this.disconnect(workspace,target);}json({ok:true});return;}

 if(route==='/leave'&&req.method==='POST'){if(member.role==='owner')throw new RequestError('请先转移所有权',403);this.database.db.prepare('DELETE FROM cloud_members WHERE workspace=? AND user=?').run(workspace,user.id);this.disconnect(workspace,user.id);json({ok:true});return;}
 if(route==='/transfer'&&req.method==='POST'){if(member.role!=='owner')throw new RequestError('需要所有者权限',403);const target=text(JSON.parse((await body(req)).toString()).user,100);if(target===user.id||!this.member(workspace,target))throw new RequestError('请选择其他现有成员');this.database.db.transaction(()=>{if(this.authorize(workspace,user.id).role!=='owner'||!this.member(workspace,target))throw new RequestError('成员或权限已变更',403);this.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run('admin',workspace,user.id);this.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run('owner',workspace,target);}).immediate();json({ok:true});return;}
 // 邀请列表包含角色与过期状态：过期口令已不可接受，但仍展示供管理员确认或重新生成。
 if(route==='/invites'&&req.method==='GET'){this.authorize(workspace,this.require(req).id,true);json(this.database.db.prepare('SELECT token AS id,email,role,expires AS expiresAt FROM cloud_invites WHERE workspace=? ORDER BY expires DESC').all(workspace).map(row=>{const invite=row as {id:string;email:string;role:string;expiresAt:number};return {...invite,expired:invite.expiresAt<=Date.now()};}));return;}
 if(route==='/invite'&&req.method==='DELETE'){this.authorize(workspace,this.require(req).id,true);const id=text(JSON.parse((await body(req)).toString()).id,64);this.authorize(workspace,this.require(req).id,true);this.database.db.prepare('DELETE FROM cloud_invites WHERE workspace=? AND token=?').run(workspace,id);json({ok:true});return;}
 if(route==='/presence'&&req.method==='POST'){const doc=text(url.searchParams.get('id'),200),input=JSON.parse((await body(req,8192)).toString());this.authorize(workspace,user.id);if(!/^doc:/.test(doc)&&doc!=='workspace:'+workspace)throw new RequestError('文档标识无效');if(!Number.isSafeInteger(input.client)||input.client<0||input.client>0xffffffff||!Number.isSafeInteger(input.clock)||input.clock<0||typeof input.active!=='boolean')throw new RequestError('在线状态无效');if(input.cursor!==null&&(!input.cursor||!position(input.cursor.anchor)||!position(input.cursor.head)))throw new RequestError('光标位置无效');const key=workspace+':'+doc+':'+input.client,old=this.presence.get(key);if(old&&old.user!==user.id)throw new RequestError('在线状态标识冲突',409);if(old&&input.clock<old.clock){json({ok:true});return;}const colors=['#3159cb','#b54708','#027a48','#7a5af8'];const state=input.active?{user:{name:user.name,color:colors[parseInt(hash(user.id).slice(0,8),16)%colors.length]},cursor:input.cursor}:null;const p={workspace,doc,client:input.client,clock:input.clock,user:user.id,state,at:Date.now()};this.presence.set(key,p);this.broadcastPresence(p);json({ok:true});return;}
 if(route==='/assets'&&req.method==='DELETE'){if(member.role!=='owner')throw new RequestError('需要所有者权限',403);this.database.db.prepare('DELETE FROM assets WHERE scope=?').run('cloud:'+workspace);json({ok:true});return;}
 if(route==='/asset'){const id=text(url.searchParams.get('id'),200),scope='cloud:'+workspace;if(req.method==='GET'){const a=this.database.db.prepare('SELECT name,mime,data FROM assets WHERE scope=? AND workspace=? AND id=?').get(scope,workspace,id) as {name:string;mime:string;data:Buffer}|undefined;if(!a){json({error:'附件不存在'},404);return;}json({id,name:a.name,mime:a.mime,size:a.data.length,data:Buffer.from(a.data).toString('base64')});return;}if(req.method==='PUT'){this.writable(workspace,req);const data=await body(req,20*1024*1024);this.writable(workspace,req);this.database.db.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?) ON CONFLICT(scope,workspace,id) DO UPDATE SET name=excluded.name,mime=excluded.mime,data=excluded.data').run(scope,workspace,id,text(url.searchParams.get('name'),1000),text(url.searchParams.get('mime'),200),data);json({ok:true});return;}if(req.method==='DELETE'){this.writable(workspace,req);this.database.db.prepare('DELETE FROM assets WHERE scope=? AND workspace=? AND id=?').run(scope,workspace,id);json({ok:true});return;}}
 const id=text(url.searchParams.get('id'),200);if(!(route==='/events'&&id==='*')&&!/^(workspace:|doc:)/.test(id))throw new RequestError('文档标识无效');const scope='cloud:'+workspace;
 if(route==='/sync'&&req.method==='POST'){const vector=await body(req,65536);this.require(req);const current=this.authorize(workspace,user.id);if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,id))throw new RequestError('文档已永久删除，请导出草稿',410);const doc=new Y.Doc();try{const bytes=this.database.load(scope,id);if(bytes)Y.applyUpdate(doc,bytes);let update:Uint8Array;try{update=Y.encodeStateAsUpdate(doc,vector);}catch{throw new RequestError('同步向量无效');}json({role:current.role,update:Buffer.from(update).toString('base64'),vector:Buffer.from(Y.encodeStateVector(doc)).toString('base64')});}finally{doc.destroy();}return;}
 if(route==='/document'&&req.method==='GET'){const bytes=this.database.load(scope,id);res.writeHead(bytes?200:204,{'Content-Type':'application/octet-stream'});res.end(bytes??undefined);return;}
 if(route==='/document'&&req.method==='POST'){this.writable(workspace,req);const bytes=await body(req,16*1024*1024);this.writable(workspace,req);if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,id))throw new RequestError('文档已永久删除，请刷新工作空间',410);this.database.append(scope,id,bytes);this.push(workspace,id,bytes);json({ok:true});return;}
 if(route==='/document'&&req.method==='DELETE'){this.writable(workspace,req);if(id.startsWith('workspace:'))throw new RequestError('请使用工作空间删除接口',403);this.database.db.transaction(()=>{this.database.db.prepare('INSERT OR IGNORE INTO cloud_deleted VALUES (?,?)').run(workspace,id);this.database.db.prepare('DELETE FROM documents WHERE scope=? AND id=?').run(scope,id);}).immediate();this.disconnect(workspace);json({ok:true});return;}
 if(route==='/events'&&req.method==='GET'){if([...this.streams.values()].filter(s=>s.user===user.id).length>=30)throw new RequestError('同步连接过多',429);res.writeHead(200,{'Content-Type':'text/event-stream','Connection':'keep-alive','X-Accel-Buffering':'no'});this.streams.set(res,{user:user.id,workspace,doc:id,expires:user.expires,session:hash(this.token(req)!),sessionId:user.session});for(const p of this.presence.values())if(p.workspace===workspace)res.write(`data: ${JSON.stringify({id:p.doc,presence:{client:p.client,clock:p.clock,state:p.state}})}\n\n`);res.write(': connected\n\n');req.on('close',()=>this.streams.delete(res));return;}
 json({error:'接口不存在'},404);
 }catch(error){if(res.headersSent){res.end();return;}json({error:error instanceof RequestError?error.message:'请求失败，请稍后重试'},error instanceof RequestError?error.status:500);}};
 close(){clearInterval(this.timer);for(const res of this.streams.keys())res.end();this.streams.clear();this.database.close();}
}
