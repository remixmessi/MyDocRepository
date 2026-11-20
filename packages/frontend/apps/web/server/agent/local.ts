import {createHash,createCipheriv,createDecipheriv,randomBytes} from 'node:crypto';
import {mkdirSync,readFileSync,writeFileSync} from 'node:fs';
import {join,dirname} from 'node:path';
import type {SQLiteProjectDatabase} from '../../../../../common/nbstore/server/database';
import {EncryptedFileSecrets,ModelConfigStore} from '../../../../../common/agent-host/src/model-config';
import {AgentService} from './service';
/** Local identity uses only its private browser Cookie, never cloud Cookies or shared preferences. */
export function createLocalAgentService(database:SQLiteProjectDatabase,worker:string,env:NodeJS.ProcessEnv=process.env){
 // Revoke credentials exposed by older clients before serving any local API.
 database.db.transaction(()=>{
  const rows=database.db.prepare("SELECT value FROM preferences WHERE key='local-account-session'").all() as {value:string}[];
  for(const row of rows)database.db.prepare('DELETE FROM local_account_sessions WHERE token=?').run(createHash('sha256').update(row.value).digest('hex'));
  database.db.prepare("DELETE FROM preferences WHERE key='local-account-session'").run();
 }).immediate();
 const directory=join(dirname(database.path),'local-agent-secrets'),keyPath=join(directory,'host.key');
 const readKey=()=>{const key=readFileSync(keyPath);if(key.length!==32)throw new Error('Invalid host encryption key');return key;};
 const secrets=new EncryptedFileSecrets(directory,{
  isEncryptionAvailable:()=>true,
  encryptString(value){mkdirSync(directory,{recursive:true,mode:0o700});try{writeFileSync(keyPath,randomBytes(32),{flag:'wx',mode:0o600});}catch(error){if((error as NodeJS.ErrnoException).code!=='EEXIST')throw error;}const iv=randomBytes(12),cipher=createCipheriv('aes-256-gcm',readKey(),iv);const data=Buffer.concat([cipher.update(value,'utf8'),cipher.final()]);return Buffer.concat([iv,cipher.getAuthTag(),data]);},
  decryptString(value){if(value.length<28)throw new Error('Invalid encrypted key');const cipher=createDecipheriv('aes-256-gcm',readKey(),value.subarray(0,12));cipher.setAuthTag(value.subarray(12,28));return Buffer.concat([cipher.update(value.subarray(28)),cipher.final()]).toString('utf8');},
 });
 return new AgentService(database,worker,req=>{
  const cookie=req.headers.cookie?.split(';').map(value=>value.trim()).find(value=>value.startsWith('mydoc_local_session='))?.slice('mydoc_local_session='.length);
  const token=cookie;if(typeof token!=='string'||!/^[a-f0-9]{64}$/.test(token))return undefined;
  const row=database.db.prepare('SELECT s.user FROM local_account_sessions s JOIN local_accounts u ON u.id=s.user WHERE s.token=? AND s.expires>?').get(createHash('sha256').update(token).digest('hex'),Date.now()) as {user:string}|undefined;
  return row?{id:row.user,scope:'local:'+row.user}:undefined;
 },env,{secrets,allowRequest:req=>{try{return ['127.0.0.1','::1','::ffff:127.0.0.1'].includes(req.socket.remoteAddress??'')&&['localhost','127.0.0.1','[::1]'].includes(new URL('http://'+req.headers.host).hostname);}catch{return false;}},loginMessage:'请先在设置中登录本地账号（与服务器账号独立）',configFor:principal=>new ModelConfigStore(database,secrets,env.MYDOC_AGENT_ALLOWED_HOSTS?.split(',').map(host=>host.trim().toLowerCase()).filter(Boolean),'agent-config:'+principal.scope),secretStorageLabel:'密钥在本机服务端加密保存；不写入数据库或备份。',disclosure:'提问将发送到你配置的模型地址。本地账号、会话和模型配置与团队服务器账号独立；自行配置的密钥仅在本机服务端加密保存，管理员环境配置保持独立。'});
}
