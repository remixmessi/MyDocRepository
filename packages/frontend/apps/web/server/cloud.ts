import {personalRestoreReceipt,restorePersonalWorkspaceArchive} from './personal-workspace-archive';
import {submitPdfForm,PdfFormError} from './office-pdf-forms';
import type {PdfFormSpec} from '../../../../common/nbstore/src/pdf-forms';
import {unpackOfficeUpdate,officeGenerationMatches} from '../../../../common/nbstore/src/office-generation';
import {submitOfficeComment,OfficeCommentError} from './office-comments';
import {assertOfficeCommentsUnchanged} from '../../../../common/nbstore/src/office-comments';
import {LOCAL_OFFICE_METHODS} from '../../../../common/nbstore/src/local-office';
import {readResourceMedia,packOfficeMedia} from './office-media';
import {officeImageMime} from '../../../../common/nbstore/src/office-image';
import {revertOfficeVersion,findOfficeVersionOperation,saveOfficeVersion,officeVersionData,OfficeVersionError} from './office-versions';
import {submitPdfOperation,PdfOperationError} from './office-pdf-operation';
import {packPdfExport} from './office-pdf-export';
import {readWorkspaceArchive,restoreLocalOfficeArchive,restoreWorkspaceArchive,workspaceRestoreReceipt,WorkspaceArchiveError} from './workspace-archive';
import {MAX_WORKSPACE_ARCHIVE_BYTES,type CloudWorkspaceArchive} from '../../../../common/nbstore/src/workspace-archive';
import {submitPresentationOperation,PresentationOperationError} from './office-presentation-operation';
import {importNativePdfAnnotations} from './office-pdf-operation';
import type {PdfNativeAnnotation} from './office-pdf-native';
import {validOfficeFocus} from '../../../../common/nbstore/src/office-presence';
import {OfficeExportQueue,OfficeExportError,officeExportPlan,runOfficeExport,captureOfficeExport,recordOfficeExport,createOfficeExportJob,listOfficeExports,readOfficeExport,deleteOfficeExport} from './office-exports';
import {FORMULA_CALCULATION_VERSION} from '../../../../common/nbstore/src/spreadsheet-formula';
import {validDelimitedImportOptions,type DelimitedImportOptions} from '../../../../common/nbstore/src/office-delimited-import';
import {OfficeRestoreError,MAX_OFFICE_BACKUP_BYTES,parseOfficeBackup,readOfficeBackup,encodeOfficeBackup,restoreOfficeBackup} from './office-backup';
import {OfficeParserPool,OfficeParserError} from './office-parser-process';
import {submitSheetOperation,SheetOperationError} from './office-sheet-operation';
import {createOfficeState,createBlankPdf} from './office-create';
import {createRequire} from 'node:module';
const cloudDatabase=(path:string,mustExist:boolean)=>{const Driver=createRequire(typeof __filename==='string'?__filename:import.meta.url)('better-sqlite3') as typeof import('better-sqlite3');return new Driver(path,{fileMustExist:mustExist});};
import {ReleaseUpdateManifests,ReleaseUpdateError,selectUpdateRelease,type UpdateReleaseRow} from './release-updates';
import * as Y from 'yjs';
import {validateOfficeTextUpdate} from './office-text';
import {submitWordTableOperation,WordTableOperationError} from './office-word-table-operation';
import type {WordTableOperation} from '../../../../common/nbstore/src/office-word-table-edit';
import {recordOfficeEdit,listOfficeActivity,insertNotifications,listNotifications,markNotificationsRead} from './office-collaboration';
import {validateSpreadsheetUpdate} from './office-spreadsheet';
import {validatePdfUpdate} from './office-pdf';
import {validatePresentationUpdate} from './office-presentation';
import {validateWordUpdate} from './office-word';
import {importOfficeOriginal,importOfficeText,OfficeImportError,OfficeTextImportError} from './office-import';
import {MAX_OFFICE_FILE_BYTES,type TextEncoding,type TextLineEnding} from '../../../../common/nbstore/src/office';
import {randomUUID,randomBytes,createHash,scryptSync,timingSafeEqual} from 'node:crypto';
import {createReadStream,createWriteStream,mkdirSync,existsSync,unlinkSync} from 'node:fs';
import {dirname,join,basename} from 'node:path';
import type {IncomingMessage,ServerResponse} from 'node:http';
import {SQLiteProjectDatabase,type SQLiteConnection} from '../../../../common/nbstore/server/database';
import {upgradeCloudDatabase} from './schema';
import {MailService} from './mail';
import {defaultAdminConfig,type DefaultAdminConfig} from './default-admin';
import {compareReleaseVersions,validReleaseVersion,type DesktopRelease} from '../../../../common/nbstore/src/releases';

const hash=(value:string)=>createHash('sha256').update(value).digest('hex');
type Account={id:string;email:string;name:string};type Member={user:string;role:'owner'|'admin'|'member'|'viewer'};
class RequestError extends Error {constructor(message:string,readonly status=400,readonly code?:string){super(message);}}
function text(value:unknown,max=160){if(typeof value!=='string'||!value.trim()||value.length>max)throw new RequestError('请输入有效内容');return value.trim();}
function compareRelease(a:{version:string;created:number},b:{version:string;created:number}){return compareReleaseVersions(b.version,a.version)||b.created-a.created;}
type ReleaseRow=Omit<DesktopRelease,'downloadPath'>&{path?:string};
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
async function bodyToFile(req:IncomingMessage,path:string,max=512*1024*1024){return await new Promise<{bytes:number;sha256:string}>((resolve,reject)=>{let bytes=0;const hash=createHash('sha256'),stream=createWriteStream(path,{flags:'wx'});const fail=(error:unknown)=>{stream.destroy();try{unlinkSync(path);}catch{}reject(error);};req.on('data',chunk=>{const data=Buffer.from(chunk);bytes+=data.length;if(bytes>max){req.destroy();fail(new RequestError('版本文件超过 512 MB',413));return;}hash.update(data);stream.write(data);});req.on('end',()=>{stream.end(()=>resolve({bytes,sha256:hash.digest('hex')}));});req.on('error',fail);stream.on('error',fail);});}
export class CloudServer {
 readonly database:SQLiteProjectDatabase;private readonly releaseDir:string;private streams=new Map<ServerResponse,{user:string;workspace:string;doc:string;expires:number;session:string;sessionId:string}>();private presence=new Map<string,{workspace:string;doc:string;client:number;clock:number;user:string;session:string;sessionId:string;state:unknown;at:number}>();private attempts=new Map<string,{count:number;until:number}>();private timer:ReturnType<typeof setInterval>;
 private localContexts=new WeakMap<IncomingMessage,{user:ReturnType<CloudServer['account']>;check:()=>void}>();
 private localOwners=new Map<string,string>();
 private ownsDatabase=true;
 private officeParser=new OfficeParserPool();private officeExports!:OfficeExportQueue;private nativeAnnotations=new Map<string,Promise<PdfNativeAnnotation[]>>();
 /** Originals are immutable, so the parsed native annotation list is cached per asset content. */
 private pdfNatives(bytes:Uint8Array){const key=createHash('sha256').update(bytes).digest('hex');let entry=this.nativeAnnotations.get(key);if(!entry){if(this.nativeAnnotations.size>=200)this.nativeAnnotations.delete(this.nativeAnnotations.keys().next().value!);entry=this.officeParser.pdfNativeAnnotations(bytes).then(value=>JSON.parse(Buffer.from(value).toString()) as PdfNativeAnnotation[]);this.nativeAnnotations.set(key,entry);entry.catch(()=>this.nativeAnnotations.delete(key));}return entry;}
 private updateManifests=new ReleaseUpdateManifests();
 readonly mail:MailService;
 constructor(path:string,options:{database?:SQLiteProjectDatabase;mail?:MailService;defaultAdmin?:DefaultAdminConfig;mustExist?:boolean;beforeMigration?:(db:SQLiteConnection)=>void}={}){this.mail=options.mail??new MailService();this.ownsDatabase=!options.database;this.database=options.database??new SQLiteProjectDatabase(path,path=>cloudDatabase(path,options.mustExist??false),upgradeCloudDatabase,{beforeMigration:options.beforeMigration,preserve:{cloud_office_versions:'workspace,resource,id,name,user,operation,created,data,sha256,deleted',cloud_users:'id,email,name,salt,password',cloud_workspaces:'id,name,created',cloud_members:'workspace,user,role',cloud_workspace_audit:'id,workspace,actor,actor_name,action,details,created'}});if(options.database&&!this.database.db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name='cloud_office_resources'").get())throw new Error('本地办公数据库尚未完成启动迁移');this.releaseDir=join(dirname(path),'releases');try{mkdirSync(this.releaseDir,{recursive:true});this.bootstrapAdmin(options.defaultAdmin);}catch(error){this.database.close();throw error;}this.officeExports=new OfficeExportQueue(this.database,this.officeParser,(workspace,resource)=>{for(const[res,stream]of this.streams)if(stream.workspace===workspace&&(stream.doc===resource||stream.doc==='*'))res.write(`data: ${JSON.stringify({id:resource,exports:true})}\n\n`);});this.officeExports.kick();this.timer=setInterval(()=>{for(const [res,state]of this.streams){if(state.expires<Date.now()||!this.database.db.prepare('SELECT 1 FROM cloud_sessions WHERE token=? AND expires>?').get(state.session,Date.now())||!this.member(state.workspace,state.user)){res.end();this.streams.delete(res);}else res.write(': heartbeat\n\n');}for(const[k,p]of this.presence)if(p.at<Date.now()-30000||!this.canReadPresence(p.workspace,p.doc,p.user)||!this.database.db.prepare('SELECT 1 FROM cloud_sessions WHERE token=? AND expires>?').get(p.session,Date.now())){this.broadcastPresence({...p,clock:p.clock+1,state:null});this.presence.delete(k);}for(const[k,v]of this.attempts)if(v.until<Date.now())this.attempts.delete(k);},15000);this.timer.unref();}
 /** Local office requests arrive only through the authenticated SQLite storage boundary.
  * Reuse the office engine without creating a cloud account, session or team workspace. */
 async localOffice(req:IncomingMessage,res:ServerResponse,scope:string,check:()=>void){
  const previous=req.url;
  try{
   check();const url=new URL(req.url??'/','http://local');const route=url.pathname.slice('/office'.length);
   if(!url.pathname.startsWith('/office/')||!LOCAL_OFFICE_METHODS[route]?.includes(req.method??''))throw new RequestError('本地办公接口不存在',404);
   const requested=text(url.searchParams.get('workspace'),100);
   const checkWorkspace=()=>{check();const data=this.database.load(scope,'registry:local'),doc=new Y.Doc();try{if(data)Y.applyUpdate(doc,data);const entry=doc.getMap<{cloud?:boolean}>('workspaces').get(requested);if(!entry||entry.cloud)throw new RequestError('本地工作区不存在',404);}finally{doc.destroy();}};
   const workspace='local-office-'+hash(JSON.stringify([scope,requested])).slice(0,64),user='local-user-'+hash(scope);
   // Workspace deletion removes its registry entry before cleaning its data. The
   // namespace is still derived from the authenticated SQLite scope, never input.
   if(route==='/office-purge'){this.database.db.transaction(()=>{check();for(const table of ['cloud_office_resources','cloud_office_versions','cloud_office_exports','cloud_notifications','cloud_office_activity','cloud_office_restores','cloud_asset_owners','cloud_deleted','cloud_document_permissions'])this.database.db.prepare(`DELETE FROM ${table} WHERE workspace=?`).run(workspace);this.database.db.prepare('DELETE FROM documents WHERE scope=?').run('cloud:'+workspace);this.database.attachments.remove('cloud:'+workspace);}).immediate();res.writeHead(200,{'Content-Type':'application/json'});res.end('{"ok":true}');return;}
   if(route==='/office-workspace-full-restore'){
    if(!/^[a-f0-9-]{36}$/i.test(requested))throw new RequestError('恢复目标无效');const operation=text(url.searchParams.get('operationId'),80);if(!/^[a-zA-Z0-9-]{16,80}$/.test(operation))throw new RequestError('恢复编号无效');const bytes=await body(req,MAX_WORKSPACE_ARCHIVE_BYTES),sha256=createHash('sha256').update(bytes).digest('hex');check();let result=personalRestoreReceipt(this.database,scope,operation,sha256);
    if(!result){const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);try{const normalized=await this.officeParser.preparePersonalArchive(bytes,requested,controller.signal);check();if(controller.signal.aborted)throw new RequestError('恢复已取消',499);result=restorePersonalWorkspaceArchive(this.database,scope,user,operation,sha256,JSON.parse(Buffer.from(normalized).toString()) as CloudWorkspaceArchive);}finally{res.off('close',cancel);}}
    const response=Buffer.from(JSON.stringify(result));res.writeHead(201,{'Content-Type':'application/json','Content-Length':response.length});res.end(response);return;
   }
   checkWorkspace();
   if(route==='/office-workspace-backup'||route==='/office-workspace-restore'){
    const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
    try{
     const input=route==='/office-workspace-backup'?readWorkspaceArchive(this.database,workspace,'个人办公文件'):await body(req,MAX_WORKSPACE_ARCHIVE_BYTES);
     const normalized=await this.officeParser.prepareWorkspaceArchive(input,workspace,controller.signal);checkWorkspace();
     if(controller.signal.aborted)throw new RequestError('备份操作已取消',499);
     if(route==='/office-workspace-restore')restoreLocalOfficeArchive(this.database,workspace,user,JSON.parse(Buffer.from(normalized).toString()) as CloudWorkspaceArchive);
     const bytes=route==='/office-workspace-backup'?normalized:Buffer.from('{"ok":true}');
     res.writeHead(200,{'Content-Type':'application/json','Content-Length':bytes.length,'X-Content-Type-Options':'nosniff'});res.end(bytes);return;
    }finally{res.off('close',cancel);}
   }
   this.localOwners.set(workspace,user);url.pathname=route;url.searchParams.set('workspace',workspace);req.url=url.pathname+url.search;
   this.localContexts.set(req,{check:checkWorkspace,user:{id:user,email:'local@localhost',name:'本机',expires:Number.MAX_SAFE_INTEGER,verified:1,serverAdmin:0,createdAt:0,avatarType:null,session:'local'}});
   if(route==='/document'){const id=text(url.searchParams.get('id'),200);if(!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id))throw new RequestError('本地办公文件不存在',404);}
   await this.handler(req,res);
  }catch(error){if(!res.headersSent){res.writeHead(error instanceof RequestError||error instanceof WorkspaceArchiveError||error instanceof OfficeParserError?error.status:400,{'Content-Type':'application/json'});res.end(JSON.stringify({error:error instanceof Error?error.message:'本地办公操作失败',...(error instanceof RequestError&&error.code?{code:error.code}:{})}));}}
  finally{this.localContexts.delete(req);req.url=previous;}
 }
 private bootstrapAdmin(config:DefaultAdminConfig|undefined){if(!config)return;
  const validated=defaultAdminConfig({MYDOC_ADMIN_EMAIL:config.email,MYDOC_ADMIN_PASSWORD:config.password,MYDOC_ADMIN_NAME:config.name})!;
  this.database.db.transaction(()=>{
   const initialized=this.database.db.prepare("SELECT user FROM cloud_bootstrap WHERE key='default-admin'").get() as {user:string}|undefined;
   if(initialized){if(!this.database.db.prepare('SELECT 1 FROM cloud_users WHERE id=? AND server_admin=1').get(initialized.user))throw new Error('默认管理员初始化记录异常，请检查数据库');return;}
   if(this.database.db.prepare('SELECT 1 FROM cloud_users WHERE email=?').get(validated.email))throw new Error('默认管理员邮箱已被现有账户使用，请配置一个未注册的邮箱');
   const id=randomUUID(),salt=randomBytes(16).toString('hex');
   this.database.db.prepare('INSERT INTO cloud_users(id,email,name,salt,password,verified,created,server_admin) VALUES (?,?,?,?,?,0,?,1)').run(id,validated.email,validated.name,salt,scryptSync(validated.password,salt,64).toString('hex'),Date.now());
   this.database.db.prepare("INSERT INTO cloud_bootstrap(key,user) VALUES ('default-admin',?)").run(id);
  }).immediate();
 }
 private member(workspace:string,user:string){if(this.localOwners.get(workspace)===user)return {user,role:'owner'} as Member;return this.database.db.prepare('SELECT user,role FROM cloud_members WHERE workspace=? AND user=?').get(workspace,user) as Member|undefined;}
 private token(req:IncomingMessage){return req.headers.cookie?.split(';').map(v=>v.trim()).find(v=>v.startsWith('mydoc_session='))?.slice(14);}
 account(req:IncomingMessage){const token=this.token(req);if(!token)return undefined;return this.database.db.prepare('SELECT u.id,u.email,u.name,u.verified,u.server_admin AS serverAdmin,u.created AS createdAt,u.avatar_type AS avatarType,s.expires,s.id AS session FROM cloud_sessions s JOIN cloud_users u ON u.id=s.user WHERE s.token=? AND s.expires>?').get(hash(token),Date.now()) as (Account&{expires:number;verified:number;serverAdmin:number;createdAt:number;avatarType:string|null;session:string})|undefined;}
 // 按 IP + 用途计数的频率限制，复用注册／登录原有的 1 分钟窗口。
 private limit(req:IncomingMessage,scope:string,max:number){const key=scope+':'+(req.socket.remoteAddress??'unknown'),limit=this.attempts.get(key)??{count:0,until:Date.now()+60000};if(limit.until<Date.now()){limit.count=0;limit.until=Date.now()+60000;}if(++limit.count>max)throw new RequestError('尝试过于频繁，请稍后重试',429);this.attempts.set(key,limit);}
 private cookie(res:ServerResponse,req:IncomingMessage,token:string){let cross=false;try{cross=typeof req.headers.origin==='string'&&new URL(req.headers.origin).hostname!==new URL('http://'+req.headers.host).hostname;}catch{cross=true;}res.setHeader('Set-Cookie',`mydoc_session=${token}; HttpOnly; SameSite=${cross?'None':'Strict'}; Path=/; Max-Age=604800${cross||req.headers['x-forwarded-proto']==='https'?' ; Secure':''}`);}
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
 private require(req:IncomingMessage){const local=this.localContexts.get(req);if(local){local.check();return local.user!;}const user=this.account(req);if(!user)throw new RequestError('请先登录',401);const expected=req.headers['x-mydoc-account']??new URL(req.url??'/','http://local').searchParams.get('account');if(expected&&expected!==user.id)throw new RequestError('账户已变更，请重新打开页面',409);return user;}
 private authorize(workspace:string,user:string,admin=false){const member=this.member(workspace,user);if(!member)throw new RequestError('你已不在此工作空间中',403);if(admin&&!['owner','admin'].includes(member.role))throw new RequestError('需要工作空间管理员权限',403);return member;}
 private writable(workspace:string,req:IncomingMessage){const user=this.require(req);const member=this.authorize(workspace,user.id);if(member.role==='viewer')throw new RequestError('查看者只能阅读工作空间',403);return member;}
 private documentAccess(workspace:string,doc:string,user:string,write=false):Member{const member=this.authorize(workspace,user);let role=member.role;if(doc.startsWith('doc:')&&!['owner','admin'].includes(role)){const permission=this.database.db.prepare('SELECT role FROM cloud_document_permissions WHERE workspace=? AND doc=? AND user=?').get(workspace,doc,user) as {role:string}|undefined;if(permission?.role==='denied')throw new RequestError('文档不存在或无权访问',404);if(permission?.role==='viewer')role='viewer';}if(write&&role==='viewer')throw new RequestError('没有该文档的编辑权限',403);return {...member,role};}
 private officeGeneration(workspace:string,id:string,supplied:unknown){
  const resource=this.database.db.prepare('SELECT generation FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {generation:string}|undefined;
  if(resource&&!officeGenerationMatches(resource.generation,supplied))throw new RequestError('文件已原位恢复，旧代际草稿已保留；请导出草稿并重新载入页面',409,'office-generation');return resource?.generation;
 }
 private officeOperation(req:IncomingMessage,workspace:string,id:string,bytes:Uint8Array){
  const input=JSON.parse(Buffer.from(bytes).toString());if(!input||typeof input!=='object'||Array.isArray(input))throw new RequestError('办公操作参数无效');
  const {officeGeneration,...operation}=input;this.officeGeneration(workspace,id,officeGeneration??req.headers['x-mydoc-office-generation']);return operation;
 }
 private permissionDocument(workspace:string,doc:string){if(!doc.startsWith('doc:')||doc.length<=4)throw new RequestError('文档标识无效');if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,doc)||!this.database.load('cloud:'+workspace,doc))throw new RequestError('文档不存在或已永久删除',404);}
 private push(workspace:string,doc:string,data:Uint8Array){for(const[res,s]of this.streams)if(s.workspace===workspace&&(s.doc===doc||s.doc==='*')){if(!res.write(`data: ${JSON.stringify({id:doc,generation:(this.database.db.prepare('SELECT generation FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,doc) as {generation:string}|undefined)?.generation,data:Buffer.from(data).toString('base64')})}\n\n`)){res.end();this.streams.delete(res);}}}
 /** Edit activity is auxiliary: a failure to record it must never fail the accepted edit. */
 private edited(workspace:string,doc:string,req:IncomingMessage){try{const user=this.require(req);recordOfficeEdit(this.database,workspace,doc,user);}catch{/* activity is best effort */}}
 private notifyUsers(users:string[]){const wanted=new Set(users);for(const[res,s]of this.streams)if(wanted.has(s.user)&&!res.write(`data: ${JSON.stringify({notifications:true})}\n\n`)){res.end();this.streams.delete(res);}}
 private canReadPresence(workspace:string,doc:string,user:string){
  try{this.documentAccess(workspace,doc,user);if(doc==='workspace:'+workspace)return true;if(!doc.startsWith('doc:')||this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,doc))return false;return !!this.database.db.prepare('SELECT 1 FROM documents WHERE scope=? AND id=?').get('cloud:'+workspace,doc);}catch{return false;}
 }
 private broadcastPresence(p:{workspace:string;doc:string;client:number;clock:number;state:unknown}){
  for(const[res,s]of this.streams)if(s.workspace===p.workspace&&(s.doc===p.doc||s.doc==='*')&&this.canReadPresence(p.workspace,p.doc,s.user)){
   if(!res.write(`data: ${JSON.stringify({id:p.doc,presence:{client:p.client,clock:p.clock,state:p.state}})}\n\n`)){res.end();this.streams.delete(res);}
  }
 }
 private clearPresence(match:(p:{workspace:string;user:string;sessionId:string})=>boolean){
  for(const[key,p]of this.presence)if(match(p)){this.broadcastPresence({...p,clock:p.clock+1,state:null});this.presence.delete(key);}
 }
 private disconnect(workspace:string,user?:string){this.clearPresence(p=>p.workspace===workspace&&(!user||p.user===user));for(const[res,s]of this.streams)if(s.workspace===workspace&&(!user||s.user===user)){res.end();this.streams.delete(res);}}
 /** 断开某账户的全部实时连接：注销、改密与重置密码后立即生效，不等心跳。 */
 private dropUserStreams(user:string,session?:string){this.clearPresence(p=>p.user===user&&(!session||p.sessionId===session));for(const[res,s]of this.streams)if(s.user===user&&(!session||s.sessionId===session)){res.end();this.streams.delete(res);}}
 private release(id:string){return this.database.db.prepare('SELECT id,version,platform,arch,filename,notes,bytes,sha256,created,path FROM cloud_releases WHERE id=?').get(id) as (ReleaseRow&{path:string})|undefined;}
 private releaseJson(row:ReleaseRow){const {path:_,...release}=row;return {...release,downloadPath:'/api/cloud/releases/download?id='+encodeURIComponent(row.id)};}
 handler=async(req:IncomingMessage,res:ServerResponse)=>{const json=(value:unknown,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(value));};try{
 res.setHeader('Cache-Control','no-store');const url=new URL(req.url??'/','http://local'),route=url.pathname;const origin=req.headers.origin;if(origin&&/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)){res.setHeader('Access-Control-Allow-Origin',origin);res.setHeader('Access-Control-Allow-Credentials','true');res.setHeader('Access-Control-Allow-Headers','Content-Type,X-Mydoc-Account,X-MyDoc-Version,X-MyDoc-Platform,X-MyDoc-Arch,X-MyDoc-Filename,X-MyDoc-Notes');res.setHeader('Access-Control-Allow-Methods','GET,POST,PATCH,PUT,DELETE,OPTIONS');if(req.method==='OPTIONS'){res.writeHead(204);res.end();return;}}else if(origin&&new URL(origin).host!==req.headers.host)throw new RequestError('禁止跨站请求',403);if(req.headers['sec-fetch-site']==='cross-site'&&!origin)throw new RequestError('禁止跨站请求',403);
 // Generic electron-updater feed, derived from existing release rows without changing their schema.
 const updateRoute=/^\/releases\/update\/(darwin|win32)\/(arm64|x64)\/(latest-mac\.yml|latest\.yml|[a-f0-9-]{36}\.(?:zip|exe))$/.exec(route);
 if(updateRoute&&['GET','HEAD'].includes(req.method??'')){
  const [,platform,arch,file]=updateRoute,extension=platform==='darwin'?'.zip':'.exe';
  if(file.startsWith('latest')){
   if(file!==(platform==='darwin'?'latest-mac.yml':'latest.yml'))throw new RequestError('更新清单不存在',404);
   const rows=this.database.db.prepare('SELECT id,version,platform,arch,filename,notes,bytes,sha256,created,path FROM cloud_releases WHERE platform=? AND arch=?').all(platform,arch) as UpdateReleaseRow[];
   try{const row=selectUpdateRelease(rows,platform,arch);const manifest=await this.updateManifests.create(row);res.writeHead(200,{'Content-Type':'application/yaml'});res.end(req.method==='HEAD'?undefined:JSON.stringify(manifest));}catch(error){if(error instanceof ReleaseUpdateError)throw new RequestError(error.message,error.status);if((error as NodeJS.ErrnoException).code==='ENOENT')throw new RequestError('更新文件不存在，请管理员重新上传',404);throw new RequestError('更新清单生成失败，请稍后重试',500);}return;
  }
  const row=this.release(file.slice(0,-4));
  if(!row||row.platform!==platform||row.arch!==arch||!file.endsWith(extension)||!row.filename.toLowerCase().endsWith(extension)||!existsSync(row.path))throw new RequestError('更新文件不存在或架构不匹配',404);
  res.setHeader('Accept-Ranges','bytes');let start=0,end=row.bytes-1,status=200;
  if(req.headers.range){const range=/^bytes=(\d+)-(\d*)$/.exec(req.headers.range);start=range?Number(range[1]):-1;end=range&&range[2]?Number(range[2]):end;if(start<0||!Number.isSafeInteger(start)||!Number.isSafeInteger(end)||end<start||start>=row.bytes){res.writeHead(416,{'Content-Range':'bytes */'+row.bytes});res.end();return;}end=Math.min(end,row.bytes-1);status=206;res.setHeader('Content-Range',`bytes ${start}-${end}/${row.bytes}`);}
  res.writeHead(status,{'Content-Type':'application/octet-stream','Content-Length':String(end-start+1)});if(req.method==='HEAD'){res.end();return;}const stream=createReadStream(row.path,{start,end});stream.on('error',()=>res.destroy());res.on('close',()=>stream.destroy());stream.pipe(res);return;
 }
 // 版本检查接口不需要登录，供桌面客户端在配置更新服务器后读取最新版本。
 if(route==='/releases/latest'&&req.method==='GET'){
  const platform=(url.searchParams.get('platform')??'').trim();if(!['darwin','win32','linux'].includes(platform))throw new RequestError('平台参数无效');
  const arch=url.searchParams.get('arch')??'';if(arch&&!['x64','arm64'].includes(arch)||platform==='darwin'&&!arch)throw new RequestError('macOS 版本检查需指定 x64 或 arm64 架构');
  const rows=this.database.db.prepare("SELECT id,version,platform,arch,filename,notes,bytes,sha256,created FROM cloud_releases WHERE platform=? AND (arch=? OR (arch='' AND platform!='darwin'))").all(platform,arch) as ReleaseRow[];const row=rows.filter(row=>validReleaseVersion(row.version)).sort(compareRelease)[0];
  json(row?this.releaseJson(row):null);return;
 }
 if(route==='/releases/download'&&req.method==='GET'){
  const row=this.release(url.searchParams.get('id')??'');if(!row||!existsSync(row.path))throw new RequestError('版本文件不存在',404);
  res.writeHead(200,{'Content-Type':'application/octet-stream','Content-Length':String(row.bytes),'Content-Disposition':`attachment; filename="${row.filename.replace(/[^\x20-\x7e]|["\\]/g,'_')}"; filename*=UTF-8''${encodeURIComponent(row.filename).replace(/['()*]/g,char=>'%'+char.charCodeAt(0).toString(16))}`});const file=createReadStream(row.path);file.on('error',()=>res.destroy());res.on('close',()=>file.destroy());file.pipe(res);return;
 }
 // 会话信息附带邮箱验证状态、到期时间与邮件服务是否可用，界面据此提示而不是猜测。
 if(route==='/session'&&req.method==='GET'){const current=this.account(req);if(current)this.database.db.prepare('UPDATE cloud_sessions SET seen=? WHERE id=?').run(Date.now(),current.session);
  json({user:current?{id:current.id,email:current.email,name:current.name,verified:!!current.verified,serverAdmin:!!current.serverAdmin,createdAt:current.createdAt,hasAvatar:!!current.avatarType,expiresAt:current.expires,session:current.session}:null,mail:{configured:this.mail.status.configured,reason:this.mail.status.reason}});return;}
 if((route==='/register'||route==='/login')&&req.method==='POST'){this.limit(req,'auth',15);const input=JSON.parse((await body(req)).toString()),email=text(input.email,254).toLowerCase(),password=text(input.password,256);if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||password.length<10)throw new RequestError('请输入有效邮箱及至少 10 位密码');let user=this.database.db.prepare('SELECT * FROM cloud_users WHERE email=?').get(email) as (Account&{salt:string;password:string;verified:number})|undefined;if(route==='/register'){if(user)throw new RequestError('无法创建账户，请尝试登录');const id=randomUUID(),name=text(input.name),salt=randomBytes(16).toString('hex');this.database.db.prepare('INSERT INTO cloud_users(id,email,name,salt,password,verified,created) VALUES (?,?,?,?,?,0,?)').run(id,email,name,salt,scryptSync(password,salt,64).toString('hex'),Date.now());user={id,email,name,salt,password:'',verified:0};}else{const salt=user?.salt??'00000000000000000000000000000000';const actual=scryptSync(password,salt,64),expected=user?Buffer.from(user.password,'hex'):Buffer.alloc(64);if(!timingSafeEqual(actual,expected)||!user)throw new RequestError('邮箱或密码不正确',401);}this.cookie(res,req,this.startSession(user.id,req));json({user:{id:user.id,email:user.email,name:user.name,verified:!!user.verified,serverAdmin:!!this.database.db.prepare('SELECT 1 FROM cloud_users WHERE id=? AND server_admin=1').get(user.id)}});return;}
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
 if(route==='/admin/releases'&&req.method==='GET'){
  if(!user.serverAdmin)throw new RequestError('需要服务器管理员权限',403);
  const rows=this.database.db.prepare('SELECT id,version,platform,arch,filename,notes,bytes,sha256,created FROM cloud_releases ORDER BY created DESC').all() as ReleaseRow[];
  json(rows.map(row=>this.releaseJson(row)));return;
 }
 if(route==='/admin/release'&&req.method==='POST'){
  if(!user.serverAdmin)throw new RequestError('需要服务器管理员权限',403);
  const version=text(req.headers['x-mydoc-version'],100);if(!validReleaseVersion(version))throw new RequestError('版本号需使用 1.2.3 或带预发布标识的格式');
  const platform=text(req.headers['x-mydoc-platform'],20);if(!['darwin','win32','linux'].includes(platform))throw new RequestError('平台无效');
  const arch=req.headers['x-mydoc-arch']??'';if(typeof arch!=='string'||arch&&!['x64','arm64'].includes(arch)||platform==='darwin'&&!arch)throw new RequestError('请选择 x64 或 arm64 架构');
  let filename:string;try{filename=basename(decodeURIComponent(text(req.headers['x-mydoc-filename'],600)));}catch{throw new RequestError('文件名无效');}let notes='';try{const encoded=typeof req.headers['x-mydoc-notes']==='string'?req.headers['x-mydoc-notes']:'';notes=Buffer.from(encoded,'base64').toString('utf8').slice(0,4000);}catch{throw new RequestError('更新说明无效');}if(!filename||filename==='.'||filename==='..')throw new RequestError('文件名无效');
  const id=randomUUID(),path=join(this.releaseDir,id+'-'+filename);const file=await bodyToFile(req,path);if(!file.bytes){unlinkSync(path);throw new RequestError('版本文件不能为空');}
  const created=Date.now();try{if(!this.require(req).serverAdmin)throw new RequestError('需要服务器管理员权限',403);this.database.db.prepare('INSERT INTO cloud_releases(id,version,platform,arch,filename,notes,bytes,sha256,created,path) VALUES (?,?,?,?,?,?,?,?,?,?)').run(id,version,platform,arch,filename,notes,file.bytes,file.sha256,created,path);}catch(error){try{unlinkSync(path);}catch{}throw error;}
  json(this.releaseJson(this.release(id)!),201);return;
 }
 if(route==='/admin/release'&&req.method==='DELETE'){
  if(!user.serverAdmin)throw new RequestError('需要服务器管理员权限',403);const id=text(JSON.parse((await body(req)).toString()).id,100),row=this.release(id);if(!row)throw new RequestError('版本不存在',404);
  this.database.db.prepare('DELETE FROM cloud_releases WHERE id=?').run(id);try{unlinkSync(row.path);}catch{}json({ok:true});return;
 }
 // 全局空间管理独立鉴权，不改变正文、附件和同步的成员校验。
 if(route.startsWith('/admin/workspace')){
  const requireAdmin=()=>{const current=this.require(req);if(!current.serverAdmin)throw new RequestError('需要服务器管理员权限',403);return current;};
  requireAdmin();const db=this.database.db;
  const page=()=>{const offset=Number(url.searchParams.get('offset')??0);if(!Number.isSafeInteger(offset)||offset<0)throw new RequestError('分页参数无效');return offset;};
  if(route==='/admin/workspaces'&&req.method==='GET'){
   const offset=page(),query=(url.searchParams.get('q')??'').trim().toLowerCase();if(query.length>160)throw new RequestError('搜索内容过长');
   const filter='instr(lower(w.name),?)>0 OR instr(lower(w.id),?)>0';
   const spaces=db.prepare(`SELECT w.id,w.name,w.created AS createdAt,(SELECT count(*) FROM cloud_members m WHERE m.workspace=w.id) AS memberCount,(SELECT u.name FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=w.id AND m.role='owner') AS ownerName,(SELECT u.email FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=w.id AND m.role='owner') AS ownerEmail FROM cloud_workspaces w WHERE ${filter} ORDER BY w.created DESC,w.id LIMIT 20 OFFSET ?`).all(query,query,offset);
   const total=(db.prepare(`SELECT count(*) total FROM cloud_workspaces w WHERE ${filter}`).get(query,query) as {total:number}).total;
   json({spaces,total,offset,limit:20});return;
  }
  if(route==='/admin/workspace-audit'&&req.method==='GET'){
   const offset=page(),workspace=url.searchParams.get('workspace');
   const filter=workspace?' WHERE workspace=?':'',params=workspace?[text(workspace,100)]:[];
   const entries=db.prepare('SELECT id,workspace,actor_name AS actorName,action,details,created AS createdAt FROM cloud_workspace_audit'+filter+' ORDER BY created DESC,id DESC LIMIT 20 OFFSET ?').all(...params,offset);
   const total=(db.prepare('SELECT count(*) total FROM cloud_workspace_audit'+filter).get(...params) as {total:number}).total;
   json({entries,total,offset,limit:20});return;
  }
  const workspace=text(url.searchParams.get('workspace'),100);
  const getSpace=()=>{const space=db.prepare('SELECT id,name,created AS createdAt FROM cloud_workspaces WHERE id=?').get(workspace) as {id:string;name:string;createdAt:number}|undefined;if(!space)throw new RequestError('工作空间不存在',404);return space;};
  if(route==='/admin/workspace'&&req.method==='GET'){
   const space=getSpace();
   const members=db.prepare('SELECT u.id,u.name,u.email,m.role FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=? ORDER BY u.name,u.id').all(workspace);
   const documents=db.prepare("SELECT count(*) count FROM documents WHERE scope=? AND id LIKE 'doc:%'").get('cloud:'+workspace) as {count:number};
   const storage=db.prepare('SELECT (SELECT coalesce(sum(length(data)),0) FROM documents WHERE scope=?) + (SELECT coalesce(sum(coalesce(f.size,length(a.data))),0) FROM assets a LEFT JOIN asset_files f ON a.scope=f.scope AND a.workspace=f.workspace AND a.id=f.id WHERE a.scope=?) AS bytes').get('cloud:'+workspace,'cloud:'+workspace) as {bytes:number};
   json({...space,members,documents:documents.count,storageBytes:storage.bytes});return;
  }
  if((route==='/admin/workspace'&&['PATCH','DELETE'].includes(req.method??''))||(route==='/admin/workspace-member'&&['PUT','PATCH','DELETE'].includes(req.method??''))||(route==='/admin/workspace-owner'&&req.method==='POST')){
   const input=JSON.parse((await body(req)).toString());let action='',details='';
   db.transaction(()=>{
    const actor=requireAdmin(),space=getSpace();
    if(route==='/admin/workspace'){
     if(req.method==='PATCH'){const name=text(input.name);db.prepare('UPDATE cloud_workspaces SET name=? WHERE id=?').run(name,workspace);action='rename';details=`空间名称：${space.name} → ${name}`;}
     else{
      if(input.confirmName!==space.name)throw new RequestError('请输入当前工作空间名称以确认删除');
      for(const table of ['cloud_members','cloud_invites','cloud_deleted','cloud_document_permissions','cloud_asset_owners','cloud_office_resources','cloud_office_restores','cloud_office_versions','cloud_office_exports','cloud_notifications','cloud_office_activity'])db.prepare(`DELETE FROM ${table} WHERE workspace=?`).run(workspace);
      db.prepare('DELETE FROM cloud_workspaces WHERE id=?').run(workspace);db.prepare('DELETE FROM documents WHERE scope=?').run('cloud:'+workspace);this.database.attachments.remove('cloud:'+workspace);
      action='delete';details=`永久删除空间：${space.name}`;
     }
    }else{
     const target=text(input.user,100),account=db.prepare('SELECT name,email FROM cloud_users WHERE id=?').get(target) as {name:string;email:string}|undefined;
     if(!account)throw new RequestError('用户不存在',404);
     const old=this.member(workspace,target),label=`${account.name}（${account.email}）`;
     if(route==='/admin/workspace-owner'){
      if(input.confirmName!==space.name)throw new RequestError('请输入当前工作空间名称以确认变更所有者');
      if(old?.role==='owner')throw new RequestError('此用户已是所有者');
      const previous=db.prepare("SELECT user FROM cloud_members WHERE workspace=? AND role='owner'").all(workspace) as {user:string}[];
      db.prepare("UPDATE cloud_members SET role='admin' WHERE workspace=? AND role='owner'").run(workspace);
      db.prepare("INSERT INTO cloud_members(workspace,user,role) VALUES (?,?,'owner') ON CONFLICT(workspace,user) DO UPDATE SET role='owner'").run(workspace,target);
      action='owner';details=`重新指定所有者：${label}；原所有者：${previous.map(row=>row.user).join('、')||'无'}`;
     }else{
      if(old?.role==='owner')throw new RequestError('请通过重新指定所有者操作变更所有者');
      if(req.method==='DELETE'){if(!old)throw new RequestError('成员不存在',404);db.prepare('DELETE FROM cloud_members WHERE workspace=? AND user=?').run(workspace,target);db.prepare('DELETE FROM cloud_document_permissions WHERE workspace=? AND user=?').run(workspace,target);action='remove';details=`移除成员：${label}（${old.role}）`;}
      else{
       if(!['admin','member','viewer'].includes(input.role))throw new RequestError('成员角色无效');
       if(req.method==='PATCH'&&!old)throw new RequestError('成员不存在',404);
       if(req.method==='PUT'&&old)throw new RequestError('用户已在空间中');
       db.prepare('INSERT INTO cloud_members(workspace,user,role) VALUES (?,?,?) ON CONFLICT(workspace,user) DO UPDATE SET role=excluded.role').run(workspace,target,input.role);
       action=old?'role':'add';details=`${old?'修改角色':'添加成员'}：${label}（${old?old.role+' → ':''}${input.role}）`;
      }
     }
    }
    db.prepare('INSERT INTO cloud_workspace_audit(id,workspace,actor,actor_name,action,details,created) VALUES (?,?,?,?,?,?,?)').run(randomUUID(),workspace,actor.id,actor.name,action,`空间“${space.name}”：${details}`,Date.now());
   }).immediate();
   this.disconnect(workspace);json({ok:true});return;
  }
  throw new RequestError('管理接口不存在',404);
 }
 if(route==='/admin/users'&&req.method==='GET'){
  if(!user.serverAdmin)throw new RequestError('需要服务器管理员权限',403);
  const offset=Number(url.searchParams.get('offset')??0);if(!Number.isSafeInteger(offset)||offset<0)throw new RequestError('分页参数无效');
  const query=(url.searchParams.get('q')??'').trim().toLowerCase();if(query.length>254)throw new RequestError('搜索内容过长');
  const workspace=url.searchParams.get('workspace');if(workspace&&!this.database.db.prepare('SELECT 1 FROM cloud_workspaces WHERE id=?').get(workspace))throw new RequestError('工作空间不存在',404);
  const filter="(instr(lower(email),?)>0 OR instr(lower(name),?)>0)"+(workspace?' AND EXISTS(SELECT 1 FROM cloud_members m WHERE m.workspace=? AND m.user=cloud_users.id)':'');
  const params=workspace?[query,query,workspace]:[query,query];
  const users=this.database.db.prepare("SELECT id,email,name,verified,created AS createdAt,server_admin AS serverAdmin,EXISTS(SELECT 1 FROM cloud_bootstrap b WHERE b.user=cloud_users.id AND b.key='default-admin') AS defaultAdmin FROM cloud_users WHERE "+filter+" ORDER BY created,id LIMIT 100 OFFSET ?").all(...params,offset);
  const total=(this.database.db.prepare('SELECT count(*) total FROM cloud_users WHERE '+filter).get(...params) as {total:number}).total;
  json({users,total,offset,limit:100});return;
 }
 if(route==='/admin/user'&&req.method==='PATCH'){
  if(!user.serverAdmin)throw new RequestError('需要服务器管理员权限',403);
  const input=JSON.parse((await body(req)).toString()),target=text(input.user,100);
  if(!this.require(req).serverAdmin)throw new RequestError('需要服务器管理员权限',403);
  if(target===user.id)throw new RequestError('不能修改当前登录账户的管理员权限');
  if(typeof input.serverAdmin!=='boolean')throw new RequestError('管理员权限参数无效');
  if(!input.serverAdmin&&this.database.db.prepare("SELECT 1 FROM cloud_bootstrap WHERE key='default-admin' AND user=?").get(target))throw new RequestError('不能取消默认管理员权限');
  const result=this.database.db.prepare('UPDATE cloud_users SET server_admin=? WHERE id=?').run(input.serverAdmin?1:0,target) as {changes:number};
  if(!result.changes)throw new RequestError('用户不存在',404);json({ok:true,serverAdmin:input.serverAdmin});return;
 }
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
 if(route==='/workspace-restore'&&req.method==='POST'){
  const operation=url.searchParams.get('operationId');if(!operation||! /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/i.test(operation))throw new RequestError('恢复编号无效');
  const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
  try{const bytes=await body(req,MAX_WORKSPACE_ARCHIVE_BYTES),sha256=createHash('sha256').update(bytes).digest('hex');if(this.require(req).id!==user.id)throw new RequestError('恢复期间账户发生变化',409);
   const receipt=workspaceRestoreReceipt(this.database,user.id,operation,sha256);if(receipt){json({...receipt,repeated:true},201);return;}
   const target=randomUUID(),normalized=await this.officeParser.prepareWorkspaceArchive(bytes,target,controller.signal);
   if(this.require(req).id!==user.id)throw new RequestError('恢复期间账户发生变化',409);if(controller.signal.aborted)throw new RequestError('恢复已取消',499);
   const archive=JSON.parse(Buffer.from(normalized).toString()) as CloudWorkspaceArchive;if(archive.workspace.id!==target)throw new RequestError('恢复目标校验失败');json(restoreWorkspaceArchive(this.database,user.id,operation,sha256,archive),201);
  }catch(error){if(error instanceof RequestError)throw error;throw new RequestError(error instanceof Error?error.message:'工作区恢复失败',error instanceof WorkspaceArchiveError||error instanceof OfficeParserError?error.status:400);}finally{res.off('close',cancel);}return;
 }
 // Notifications span workspaces; each row is re-checked against current membership and file access when listed.
 if(route==='/notifications'&&req.method==='GET'){const limit=Number(url.searchParams.get('limit')??50);json(listNotifications(this.database,user.id,(space,resource)=>!!this.member(space,user.id)&&this.canReadPresence(space,resource,user.id),Number.isSafeInteger(limit)?Math.max(1,Math.min(limit,200)):50));return;}
 if(route==='/notifications-read'&&req.method==='POST'){const input=JSON.parse((await body(req,64*1024)).toString()) as {ids?:unknown};if(input.ids!==undefined&&(!Array.isArray(input.ids)||input.ids.length>500||input.ids.some(id=>typeof id!=='string'||id.length>100)))throw new RequestError('通知编号无效');json({changed:markNotificationsRead(this.database,user.id,input.ids as string[]|undefined)});this.notifyUsers([user.id]);return;}
 const workspace=text(url.searchParams.get('workspace'),100);let member=this.authorize(workspace,user.id);
 if(route==='/office-mentionable'&&req.method==='GET'){const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);json((this.database.db.prepare('SELECT u.id,u.name FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=? AND u.id<>? ORDER BY u.name').all(workspace,user.id) as {id:string;name:string}[]).filter(row=>this.canReadPresence(workspace,id,row.id)));return;}
 if(route==='/office-activity'&&req.method==='GET'){const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);if(!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id))throw new RequestError('办公文件不存在或已删除',404);json(listOfficeActivity(this.database,workspace,id));return;}
 if(route==='/workspace-backup'&&req.method==='GET'){
  if(member.role!=='owner')throw new RequestError('完整备份需要工作区所有者权限',403);const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
  try{const snapshot=readWorkspaceArchive(this.database,workspace),bytes=await this.officeParser.prepareWorkspaceArchive(snapshot,undefined,controller.signal),current=this.require(req);
   if(current.id!==user.id||this.authorize(workspace,current.id).role!=='owner')throw new RequestError('备份期间账户或权限发生变化',403);if(controller.signal.aborted)throw new RequestError('备份已取消',499);
   const name=(this.database.db.prepare('SELECT name FROM cloud_workspaces WHERE id=?').get(workspace) as {name:string}|undefined)?.name;if(!name)throw new RequestError('工作区已删除',404);
   res.writeHead(200,{'Content-Type':'application/json','Content-Length':bytes.length,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(name+'.mydoc-team.json'),'X-Content-Type-Options':'nosniff'});res.end(bytes);
  }catch(error){if(error instanceof RequestError)throw error;throw new RequestError(error instanceof Error?error.message:'工作区备份失败',error instanceof WorkspaceArchiveError||error instanceof OfficeParserError?error.status:400);}finally{res.off('close',cancel);}return;
 }

 if(route==='/access'&&req.method==='GET'){const doc=url.searchParams.get('id');json({role:doc?this.documentAccess(workspace,doc,user.id).role:member.role});return;}
 // 空间概览：名称、所有者、当前角色、成员数量与存储占用，供设置页集中展示。
 if(route==='/overview'&&req.method==='GET'){const space=this.database.db.prepare('SELECT name,created FROM cloud_workspaces WHERE id=?').get(workspace) as {name:string;created:number};const owner=this.database.db.prepare("SELECT u.name,u.email FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=? AND m.role='owner'").get(workspace) as {name:string;email:string}|undefined;const counts=this.database.db.prepare("SELECT count(*) total,sum(role='owner') owners,sum(role='admin') admins,sum(role='member') members,sum(role='viewer') viewers FROM cloud_members WHERE workspace=?").get(workspace) as {total:number;owners:number;admins:number;members:number;viewers:number};const documents=this.database.db.prepare("SELECT count(*) count,coalesce(sum(length(data)),0) bytes FROM documents WHERE scope=? AND id LIKE 'doc:%'").get('cloud:'+workspace) as {count:number;bytes:number};const shared=this.database.db.prepare("SELECT coalesce(sum(length(data)),0) bytes FROM documents WHERE scope=? AND id LIKE 'workspace:%'").get('cloud:'+workspace) as {bytes:number};const assets=this.database.db.prepare('SELECT count(*) count,coalesce(sum(coalesce(f.size,length(a.data))),0) bytes FROM assets a LEFT JOIN asset_files f ON a.scope=f.scope AND a.workspace=f.workspace AND a.id=f.id WHERE a.scope=?').get('cloud:'+workspace) as {count:number;bytes:number};const invites=this.database.db.prepare('SELECT count(*) count FROM cloud_invites WHERE workspace=? AND expires>?').get(workspace,Date.now()) as {count:number};json({id:workspace,name:space.name,createdAt:space.created,role:member.role,owner:owner?{name:owner.name,email:owner.email}:null,members:counts,documents:documents.count,invites:invites.count,storage:{documents:documents.bytes+shared.bytes,assets:assets.bytes,assetCount:assets.count,total:documents.bytes+shared.bytes+assets.bytes}});return;}
 if(route==='/workspace'&&req.method==='PATCH'){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString());this.authorize(workspace,this.require(req).id,true);this.database.db.prepare('UPDATE cloud_workspaces SET name=? WHERE id=?').run(text(input.name),workspace);json({ok:true});return;}
 if(route==='/workspace'&&req.method==='DELETE'){if(member.role!=='owner')throw new RequestError('只有所有者可以删除工作空间',403);this.database.db.transaction(()=>{for(const table of ['cloud_members','cloud_invites','cloud_deleted','cloud_document_permissions','cloud_asset_owners','cloud_office_resources','cloud_office_restores','cloud_office_versions','cloud_office_exports','cloud_notifications','cloud_office_activity'])this.database.db.prepare(`DELETE FROM ${table} WHERE workspace=?`).run(workspace);this.database.db.prepare('DELETE FROM cloud_workspaces WHERE id=?').run(workspace);this.database.db.prepare('DELETE FROM documents WHERE scope=?').run('cloud:'+workspace);this.database.attachments.remove('cloud:'+workspace);}).immediate();this.disconnect(workspace);json({ok:true});return;}
 if(route==='/members'&&req.method==='GET'){json(this.database.db.prepare('SELECT u.id,u.name,u.email,m.role FROM cloud_members m JOIN cloud_users u ON u.id=m.user WHERE m.workspace=?').all(workspace));return;}
 if(route==='/directory-users'&&req.method==='GET'){
  this.authorize(workspace,user.id,true);const query=(url.searchParams.get('q')??'').trim().toLowerCase();if(query.length>254)throw new RequestError('搜索内容过长');
  json(this.database.db.prepare("SELECT u.id,u.name,u.email FROM cloud_users u WHERE (instr(lower(u.email),?)>0 OR instr(lower(u.name),?)>0) AND NOT EXISTS (SELECT 1 FROM cloud_members m WHERE m.workspace=? AND m.user=u.id) ORDER BY u.name,u.email LIMIT 50").all(query,query,workspace));return;
 }
 // 办公资源目录：文件原件仍通过 /asset 保存，目录与协作状态通过独立资源表管理。
 if(route==='/office-deleted-resources'&&req.method==='GET'){this.authorize(workspace,user.id,true);const offset=Number(url.searchParams.get('offset')??0);if(!Number.isSafeInteger(offset)||offset<0||offset>1000000)throw new RequestError('分页位置无效');const rows=this.database.db.prepare('SELECT r.id,r.name,r.updated,(SELECT count(*) FROM cloud_office_versions v WHERE v.workspace=r.workspace AND v.resource=r.id AND v.deleted=0) AS versions FROM cloud_office_resources r WHERE r.workspace=? AND r.deleted=1 ORDER BY r.updated DESC,r.id LIMIT 101 OFFSET ?').all(workspace,offset);json({files:rows.slice(0,100),more:rows.length>100});return;}
 if(['/office-versions','/office-version','/office-version-restore','/office-version-revert'].includes(route)){
  const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);const source=()=>{const row=this.database.db.prepare('SELECT deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {deleted:number}|undefined;if(!row||row.deleted&&!['owner','admin'].includes(this.authorize(workspace,user.id).role))throw new RequestError('办公文件不存在或已删除',404);return row;};source();
  const assertCurrent=(write=false)=>{const current=this.require(req);if(current.id!==user.id)throw new RequestError('处理期间账户发生变化',409);this.documentAccess(workspace,id,current.id,write);if(write)this.writable(workspace,req);source();return current;};
  const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
  try{
   if(route==='/office-versions'&&req.method==='GET'){json((this.database.db.prepare("SELECT v.id,v.name,v.created,length(v.data) AS bytes,v.sha256,v.user,coalesce(u.name,'原成员') AS author FROM cloud_office_versions v LEFT JOIN cloud_users u ON u.id=v.user WHERE v.workspace=? AND v.resource=? AND v.deleted=0 ORDER BY v.created DESC,v.id LIMIT 100").all(workspace,id) as {user:string}[]).map(row=>({...row,canDelete:row.user===user.id||['owner','admin'].includes(member.role)})));return;}
   if(route==='/office-versions'&&req.method==='POST'){
    assertCurrent(true);if(source().deleted)throw new RequestError('已删除文件不能保存新的版本',410);const input=JSON.parse((await body(req)).toString()),name=text(input.name,120),operation=text(input.operationId,80);if(!/^[a-zA-Z0-9-]{16,80}$/.test(operation)||/[\x00-\x1f]/.test(name))throw new RequestError('版本名称或操作编号无效');assertCurrent(true);
    const receipt=findOfficeVersionOperation(this.database,workspace,id,user.id,operation,name);if(receipt){json({id:receipt.id,repeated:true},201);return;}
    const snapshot=readOfficeBackup(this.database,workspace,id),state=await this.officeParser.validateBackup(snapshot.archive.resource.kind,packOfficeMedia(snapshot.state,snapshot.archive.media??[]),controller.signal);assertCurrent(true);if(controller.signal.aborted)throw new RequestError('保存版本已取消',499);
    json(saveOfficeVersion(this.database,workspace,id,user.id,operation,name,encodeOfficeBackup(snapshot.archive,state)),201);return;
   }
   const version=text(url.searchParams.get('version'),100);
   if(route==='/office-version'&&req.method==='DELETE'){
    assertCurrent(true);const row=this.database.db.prepare('SELECT user FROM cloud_office_versions WHERE workspace=? AND resource=? AND id=?').get(workspace,id,version) as {user:string}|undefined;if(!row)throw new RequestError('版本不存在',404);const role=this.authorize(workspace,user.id).role;if(row.user!==user.id&&!['owner','admin'].includes(role))throw new RequestError('只能删除自己保存的版本，管理员可管理全部版本',403);
    this.database.db.prepare('UPDATE cloud_office_versions SET deleted=1,data=NULL WHERE workspace=? AND resource=? AND id=?').run(workspace,id,version);json({ok:true});return;
   }
   const snapshot=officeVersionData(this.database,workspace,id,version);
   if(route==='/office-version'&&req.method==='GET'){res.writeHead(200,{'Content-Type':'application/json','Content-Length':snapshot.data.length,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(snapshot.name+'.mydoc-office.json'),'X-Content-Type-Options':'nosniff','X-MyDoc-Snapshot':snapshot.sha256});res.end(snapshot.data);return;}
   if(route==='/office-version-revert'&&req.method==='POST'){
    assertCurrent(true);if(source().deleted)throw new RequestError('已删除文件不能原位恢复',410);const input=JSON.parse((await body(req,4096)).toString());
    if(!input||! /^[a-zA-Z0-9-]{16,60}$/.test(input.operationId??'')||typeof input.generation!=='string'||input.generation.length>100||! /^[a-f0-9]{64}$/.test(input.base??''))throw new RequestError('恢复编号、代际或当前快照无效');
    const parsed=parseOfficeBackup(snapshot.data),state=await this.officeParser.validateBackup(parsed.archive.resource.kind,packOfficeMedia(parsed.state,parsed.archive.media??[]),controller.signal);assertCurrent(true);officeVersionData(this.database,workspace,id,version);if(controller.signal.aborted)throw new RequestError('恢复已取消',499);
    const result=revertOfficeVersion(this.database,workspace,id,user.id,{...input,version},parsed.archive,state);if(!result.repeated){const current=this.database.load('cloud:'+workspace,id)!;this.push(workspace,id,current);this.edited(workspace,id,req);this.disconnect(workspace);}json(result);return;
   }
   if(route==='/office-version-restore'&&req.method==='POST'){
    assertCurrent(true);const input=JSON.parse((await body(req)).toString()),operation=text(input.operationId,80);if(!/^[a-zA-Z0-9-]{16,80}$/.test(operation))throw new RequestError('恢复编号无效');const parsed=parseOfficeBackup(snapshot.data),state=await this.officeParser.validateBackup(parsed.archive.resource.kind,packOfficeMedia(parsed.state,parsed.archive.media??[]),controller.signal);const suffix='（'+snapshot.name+'）.'+parsed.archive.resource.sourceFormat;parsed.archive.resource.name=parsed.archive.resource.name.replace(/\.[^.]+$/,'').slice(0,240-suffix.length)+suffix;assertCurrent(true);officeVersionData(this.database,workspace,id,version);if(controller.signal.aborted)throw new RequestError('恢复已取消',499);
    const restored=this.database.db.transaction(()=>{const existing=this.database.db.prepare('SELECT 1 FROM cloud_office_restores WHERE workspace=? AND user=? AND operation=?').get(workspace,user.id,operation);const resource=restoreOfficeBackup(this.database,workspace,user.id,parsed,state,{id:operation,sha256:createHash('sha256').update(JSON.stringify([id,version,snapshot.sha256])).digest('hex')});if(!existing)this.database.db.prepare('INSERT OR IGNORE INTO cloud_document_permissions(workspace,doc,user,role) SELECT workspace,?,user,role FROM cloud_document_permissions WHERE workspace=? AND doc=?').run(resource.id,workspace,id);return resource;}).immediate();json(restored,201);return;
   }
   throw new RequestError('版本操作不支持',405);
  }catch(error){if(error instanceof RequestError)throw error;throw new RequestError(error instanceof Error?error.message:'版本操作失败',error instanceof OfficeVersionError||error instanceof OfficeParserError||error instanceof OfficeRestoreError?error.status:400);}finally{res.off('close',cancel);}
 }
 if((route==='/office-backup'&&req.method==='GET')||(route==='/office-restore'&&req.method==='POST')){
  const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
  try{
   if(route==='/office-backup'){
    const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);if(!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id))throw new RequestError('办公文件不存在或已删除',404);const snapshot=readOfficeBackup(this.database,workspace,id);
    const state=await this.officeParser.validateBackup(snapshot.archive.resource.kind,packOfficeMedia(snapshot.state,snapshot.archive.media??[]),controller.signal);
    const current=this.require(req);this.documentAccess(workspace,id,current.id);if(current.id!==user.id||!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id))throw new RequestError('文件或权限已改变',404);
    const bytes=encodeOfficeBackup(snapshot.archive,state);res.writeHead(200,{'Content-Type':'application/json','Content-Length':bytes.length,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(snapshot.archive.resource.name+'.mydoc-office.json'),'X-Content-Type-Options':'nosniff'});res.end(bytes);
   }else{
    this.writable(workspace,req);const operationId=url.searchParams.get('operationId');if(operationId!==null&&!/^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/i.test(operationId))throw new RequestError('恢复编号无效');const bytes=await body(req,MAX_OFFICE_BACKUP_BYTES);const parsed=parseOfficeBackup(bytes);const state=await this.officeParser.validateBackup(parsed.archive.resource.kind,packOfficeMedia(parsed.state,parsed.archive.media??[]),controller.signal);
    const current=this.require(req);if(current.id!==user.id)throw new RequestError('恢复期间账户发生变化',409);this.writable(workspace,req);if(controller.signal.aborted)throw new RequestError('恢复已取消',499);
    const resource=restoreOfficeBackup(this.database,workspace,current.id,parsed,state,operationId?{id:operationId,sha256:createHash('sha256').update(bytes).digest('hex')}:undefined);this.documentAccess(workspace,resource.id,current.id);json(resource,201);
   }
  }catch(error){if(error instanceof RequestError)throw error;throw new RequestError(error instanceof Error?error.message:'办公文件备份失败',error instanceof OfficeParserError||error instanceof OfficeRestoreError?error.status:400);}finally{res.off('close',cancel);}
  return;
 }
 if(route==='/office-text-import'&&req.method==='POST'){
  this.writable(workspace,req);
  const name=text(url.searchParams.get('name'),240),encoding=url.searchParams.get('encoding')??undefined,lineEnding=url.searchParams.get('lineEnding')??undefined;
  if(!/\.txt$/i.test(name))throw new RequestError('请选择 TXT 文件');
  if(encoding&&!['utf-8','utf-16le','utf-16be'].includes(encoding)||lineEnding&&!['lf','crlf','cr'].includes(lineEnding))throw new RequestError('文本导入设置无效');
  const bytes=await body(req,MAX_OFFICE_FILE_BYTES);this.writable(workspace,req);const current=this.require(req);
  let resource;try{resource=importOfficeText(this.database,workspace,current.id,name,bytes,encoding as TextEncoding|undefined,lineEnding as TextLineEnding|undefined);}catch(error){if(error instanceof OfficeTextImportError)throw new RequestError(error.message);throw error;}
  json(resource,201);return;
 }
 if(route==='/office-import'&&req.method==='POST'){
  this.writable(workspace,req);
  const name=text(url.searchParams.get('name'),240);
  const delimited=/\.(csv|tsv)$/i.test(name),encoding=url.searchParams.get('encoding'),delimiter=url.searchParams.get('delimiter'),quoteMode=url.searchParams.get('quoteMode');
  if(!delimited&&(encoding!==null||delimiter!==null||quoteMode!==null))throw new RequestError('这些导入设置仅适用于 CSV/TSV');
  const importOptions={...(encoding===null?{}:{encoding}),delimiter:delimiter===null?(/\.tsv$/i.test(name)?'\t':','):delimiter==='tab'?'\t':delimiter==='comma'?',':delimiter,quoteMode:quoteMode??'csv'};
  if(delimited&&!validDelimitedImportOptions(importOptions))throw new RequestError('CSV/TSV 导入设置无效');
  const bytes=await body(req,MAX_OFFICE_FILE_BYTES);this.writable(workspace,req);const current=this.require(req);
  const controller=new AbortController(),onClose=()=>{if(!res.writableEnded)controller.abort();};res.once('close',onClose);
  try{
   const resource=await importOfficeOriginal(this.database,workspace,current.id,name,bytes,(format,input)=>format==='csv'||format==='tsv'?this.officeParser.parseDelimited(input,importOptions as DelimitedImportOptions,controller.signal):this.officeParser.parse(format,input,controller.signal),()=>{
    if(res.destroyed||controller.signal.aborted)throw new RequestError('文件导入已取消',499);
    this.writable(workspace,req);if(this.require(req).id!==current.id)throw new RequestError('导入期间账户发生变化',409);
   });
   json(resource,201);
  }catch(error){if(error instanceof OfficeImportError)throw new RequestError(error.message,error.status);throw error;}
  finally{res.off('close',onClose);}
  return;
 }
 if(route==='/office-export'&&req.method==='GET'){
  const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);
  let plan,captured;try{plan=officeExportPlan(url.searchParams.get('format'),url.searchParams.get('scope'),text(url.searchParams.get('sheet')??'first',128));captured=captureOfficeExport(this.database,workspace,id,plan);}catch(error){if(error instanceof OfficeExportError)throw new RequestError(error.message,error.status);throw error;}
  const controller=new AbortController(),onClose=()=>{if(!res.writableEnded)controller.abort();};res.once('close',onClose);
  try{
   const bytes=await runOfficeExport(this.officeParser,plan,captured.input,controller.signal);
   if(res.destroyed||controller.signal.aborted)throw new RequestError('文件导出已取消',499);
   const current=this.require(req);if(current.id!==user.id)throw new RequestError('导出期间账户发生变化',409);
   this.documentAccess(workspace,id,current.id);
   const resource=this.database.db.prepare('SELECT deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {deleted:number}|undefined;
   if(!resource||resource.deleted)throw new RequestError('办公文件不存在或已删除',404);
   // History is best effort: a failure to record must not lose the download the user is waiting for.
   try{recordOfficeExport(this.database,workspace,id,current.id,plan,captured,bytes);}catch{}
   res.writeHead(200,{'Content-Type':plan.mime,'Content-Length':bytes.length,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(captured.filename),'X-MyDoc-Snapshot':captured.hash,...(plan.calculated?{'X-MyDoc-Calculation-Version':FORMULA_CALCULATION_VERSION}:{}),'X-Content-Type-Options':'nosniff','Cache-Control':'no-store'});res.end(bytes);
  }catch(error){if(error instanceof OfficeParserError)throw new RequestError(error.message,error.status);throw error;}
  finally{res.off('close',onClose);}
  return;
 }
 if(route==='/office-exports'&&['GET','POST','DELETE'].includes(req.method??'')){
  const id=text(url.searchParams.get('id'),200),write=req.method!=='GET';const access=this.documentAccess(workspace,id,user.id);if(write)this.writable(workspace,req);
  if(!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id))throw new RequestError('办公文件不存在或已删除',404);
  try{
   if(req.method==='GET'){const job=url.searchParams.get('job');if(!job){json(listOfficeExports(this.database,workspace,id));return;}
    const output=readOfficeExport(this.database,workspace,id,text(job,64));res.writeHead(200,{'Content-Type':output.mime,'Content-Length':output.output.length,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(output.filename),'X-MyDoc-Snapshot':output.snapshot,'X-MyDoc-Export-Sha256':output.sha256,'X-Content-Type-Options':'nosniff','Cache-Control':'no-store'});res.end(output.output);return;}
   if(req.method==='DELETE'){deleteOfficeExport(this.database,workspace,id,text(url.searchParams.get('job'),64),user.id,['owner','admin'].includes(access.role));json({ok:true});return;}
   const input=JSON.parse(Buffer.from(await body(req,4096)).toString()) as {operationId?:unknown;format?:unknown;scope?:unknown;sheet?:unknown};if(!input||typeof input!=='object'||Object.keys(input).some(key=>!['operationId','format','scope','sheet','officeGeneration'].includes(key)))throw new RequestError('导出任务参数无效');
   this.documentAccess(workspace,id,this.require(req).id);const plan=officeExportPlan(typeof input.format==='string'?input.format:null,typeof input.scope==='string'?input.scope:null,typeof input.sheet==='string'?input.sheet:'first');
   const result=createOfficeExportJob(this.database,workspace,id,user.id,String(input.operationId??''),plan);this.officeExports.kick();json(result,result.repeated?200:202);return;
  }catch(error){if(error instanceof OfficeExportError)throw new RequestError(error.message,error.status);throw error;}
 }
 if(route==='/office-original'&&req.method==='GET'){
  const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);
  const resource=this.database.db.prepare('SELECT original_asset_id,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {original_asset_id:string|null;deleted:number}|undefined;
  if(!resource||resource.deleted||!resource.original_asset_id)throw new RequestError('办公原件不存在或已删除',404);
  const asset=this.database.attachments.get('cloud:'+workspace,workspace,resource.original_asset_id);if(!asset)throw new RequestError('办公原件不存在',404);
  res.writeHead(200,{'Content-Type':'application/octet-stream','Content-Length':asset.data.length,'Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(asset.name),'X-Content-Type-Options':'nosniff','Cache-Control':'no-store'});res.end(asset.data);return;
 }
 if(route==='/office-comment-operation'&&req.method==='POST'){
  this.writable(workspace,req);const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id,true);const input=this.officeOperation(req,workspace,id,await body(req,16384));const current=this.require(req);if(current.id!==user.id)throw new RequestError('评论期间账户变化',409);this.writable(workspace,req);const access=this.documentAccess(workspace,id,current.id,true);
  if(!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id))throw new RequestError('办公文件不存在或已删除',404);
  const mentions=(input as {mentions?:unknown}).mentions;if(Array.isArray(mentions)&&mentions.some(target=>typeof target!=='string'||target===current.id||!this.member(workspace,target)||!this.canReadPresence(workspace,id,target)))throw new RequestError('只能提及可以访问此文件的其他成员');
  // Comment state and its notifications commit together, so a retried operation id never notifies twice.
  try{const {result,notified}=this.database.db.transaction(()=>{const result=submitOfficeComment(this.database,workspace,id,{id:current.id,name:current.name,admin:['owner','admin'].includes(access.role)},input);return {result,notified:insertNotifications(this.database,workspace,id,current,result.notify.filter(notice=>this.member(workspace,notice.user)&&this.canReadPresence(workspace,id,notice.user)))};}).immediate();
   if(!result.repeated)this.push(workspace,id,result.update);this.notifyUsers(notified);json({update:Buffer.from(result.update).toString('base64'),repeated:result.repeated});}catch(error){if(error instanceof OfficeCommentError)throw new RequestError(error.message,error.status);throw error;}return;
 }
 if((route==='/office-pdf-forms'&&req.method==='GET')||(route==='/office-pdf-form-operation'&&req.method==='POST')){
  const id=text(url.searchParams.get('id'),200),write=req.method==='POST';this.documentAccess(workspace,id,user.id,write);if(write)this.writable(workspace,req);const resource=()=>this.database.db.prepare("SELECT original_asset_id AS asset,generation FROM cloud_office_resources WHERE workspace=? AND id=? AND kind='pdf' AND deleted=0").get(workspace,id) as {asset:string;generation:string}|undefined,original=resource();if(!original)throw new RequestError('PDF 不存在',404);
  const input=write?this.officeOperation(req,workspace,id,await body(req,100*1024)):undefined,asset=this.database.attachments.get('cloud:'+workspace,workspace,original.asset);if(!asset)throw new RequestError('PDF 原件缺失',404);const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
  try{const specs=JSON.parse(Buffer.from(await this.officeParser.pdfFormSpecs(asset.data,controller.signal)).toString()) as PdfFormSpec[];if(this.require(req).id!==user.id)throw new RequestError('处理期间账户发生变化',409);this.documentAccess(workspace,id,user.id,write);const latest=resource();if(!latest||latest.asset!==original.asset)throw new RequestError('PDF 已删除或原件已变化',409);if(write&&latest.generation!==original.generation)throw new RequestError('文件已原位恢复，请重新载入',409,'office-generation');if(controller.signal.aborted)throw new RequestError('表单操作已取消',499);
   if(write){this.writable(workspace,req);const result=submitPdfForm(this.database,workspace,id,user.id,input as import('../../../../common/nbstore/src/pdf-forms').PdfFormOperation,specs);if(!result.repeated){this.push(workspace,id,result.update);this.edited(workspace,id,req);}json({...result,update:Buffer.from(result.update).toString('base64')});}else json(specs);
  }catch(error){if(error instanceof OfficeParserError||error instanceof PdfFormError)throw new RequestError(error.message,error.status);throw error;}finally{res.off('close',cancel);}return;
 }
 if(route==='/office-pdf-native-annotations'&&['GET','POST'].includes(req.method??'')){
  const id=text(url.searchParams.get('id'),200),write=req.method==='POST';this.documentAccess(workspace,id,user.id,write);if(write)this.writable(workspace,req);const resource=()=>this.database.db.prepare("SELECT original_asset_id AS asset,generation FROM cloud_office_resources WHERE workspace=? AND id=? AND kind='pdf' AND deleted=0").get(workspace,id) as {asset:string;generation:string}|undefined,original=resource();if(!original)throw new RequestError('PDF 不存在',404);
  const input=write?this.officeOperation(req,workspace,id,await body(req,4096)) as {operationId?:unknown}:undefined;if(write&&(Object.keys(input!).length!==1||typeof input!.operationId!=='string'))throw new RequestError('原件批注导入参数无效');
  const asset=this.database.attachments.get('cloud:'+workspace,workspace,original.asset);if(!asset)throw new RequestError('PDF 原件缺失',404);const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);
  try{const natives=await this.pdfNatives(asset.data);if(this.require(req).id!==user.id)throw new RequestError('处理期间账户发生变化',409);this.documentAccess(workspace,id,user.id,write);const latest=resource();if(!latest||latest.asset!==original.asset)throw new RequestError('PDF 已删除或原件已变化',409);if(write&&latest.generation!==original.generation)throw new RequestError('文件已原位恢复，请重新载入',409,'office-generation');if(controller.signal.aborted)throw new RequestError('导入已取消',499);
   if(write){this.writable(workspace,req);const result=importNativePdfAnnotations(this.database,workspace,id,user.id,input!.operationId as string,natives);if(!result.repeated){this.push(workspace,id,result.update);this.edited(workspace,id,req);}json({...result,update:Buffer.from(result.update).toString('base64')});}
   else json(natives.map(({key,page,subtype,reason,annotation})=>({key,page,subtype,importable:!!annotation,...(reason?{reason}:{})})));
  }catch(error){if(error instanceof OfficeParserError||error instanceof PdfOperationError)throw new RequestError(error.message,error.status);throw error;}finally{res.off('close',cancel);}return;
 }
 if(route==='/office-pdf-operation'&&req.method==='POST'){
  this.writable(workspace,req);const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id,true);const input=this.officeOperation(req,workspace,id,await body(req,256*1024));const current=this.require(req);this.writable(workspace,req);this.documentAccess(workspace,id,current.id,true);
  const resource=this.database.db.prepare('SELECT kind,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {kind:string;deleted:number}|undefined;if(!resource||resource.deleted||resource.kind!=='pdf')throw new RequestError('PDF 不存在',404);
  try{const result=submitPdfOperation(this.database,workspace,id,current.id,input);if(!result.repeated){this.push(workspace,id,result.update);this.edited(workspace,id,req);}json({...result,update:Buffer.from(result.update).toString('base64')});}catch(error){if(error instanceof PdfOperationError)throw new RequestError(error.message,error.status);throw error;}return;
 }
 if(route==='/office-pdf-export'&&req.method==='GET'){
  const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id);const resource=this.database.db.prepare("SELECT name,original_asset_id FROM cloud_office_resources WHERE workspace=? AND id=? AND kind='pdf' AND deleted=0").get(workspace,id) as {name:string;original_asset_id:string}|undefined;
  if(!resource)throw new RequestError('PDF 不存在',404);const original=this.database.attachments.get('cloud:'+workspace,workspace,resource.original_asset_id),state=this.database.load('cloud:'+workspace,id);if(!original||!state)throw new RequestError('PDF 原件或协作状态缺失',404);
  const controller=new AbortController(),cancel=()=>controller.abort();res.on('close',cancel);try{const flatten=url.searchParams.get('flatten')==='1',bytes=await this.officeParser.exportAnnotatedPdf(packPdfExport(state,original.data),controller.signal,flatten);if(this.require(req).id!==user.id)throw new RequestError('导出期间账户变化',409);this.documentAccess(workspace,id,user.id);if(!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,id))throw new RequestError('PDF 已删除',404);
   res.writeHead(200,{'Content-Type':'application/pdf','Content-Disposition':"attachment; filename*=UTF-8''"+encodeURIComponent(resource.name.replace(/\.pdf$/i,'')+(flatten?'-扁平化副本.pdf':'-批注副本.pdf')),'X-MyDoc-Snapshot':createHash('sha256').update(state).update(original.data).digest('hex'),'X-Content-Type-Options':'nosniff'});res.end(bytes);
  }catch(error){if(error instanceof OfficeParserError)throw new RequestError(error.message,error.status);throw error;}finally{res.off('close',cancel);}return;
 }
 if(route==='/office-presentation-operation'&&req.method==='POST'){
  this.writable(workspace,req);const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id,true);const input=this.officeOperation(req,workspace,id,await body(req,1024*1024));this.writable(workspace,req);const currentUser=this.require(req);this.documentAccess(workspace,id,currentUser.id,true);
  const resource=this.database.db.prepare('SELECT kind,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {kind:string;deleted:number}|undefined;if(!resource||resource.deleted||resource.kind!=='presentation')throw new RequestError('演示文稿不存在',404);
  let result;try{result=submitPresentationOperation(this.database,workspace,id,currentUser.id,input);}catch(error){if(error instanceof PresentationOperationError)throw new RequestError(error.message,error.status);throw error;}
  if(!result.repeated){this.push(workspace,id,result.update);this.edited(workspace,id,req);}json({...result,update:Buffer.from(result.update).toString('base64')});return;
 }
 if(route==='/office-word-table-operation'&&req.method==='POST'){
  this.writable(workspace,req);const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id,true);const input=this.officeOperation(req,workspace,id,await body(req,4096));this.writable(workspace,req);const currentUser=this.require(req);this.documentAccess(workspace,id,currentUser.id,true);
  const resource=this.database.db.prepare('SELECT kind,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {kind:string;deleted:number}|undefined;if(!resource||resource.deleted||resource.kind!=='word')throw new RequestError('Word 文档不存在',404);
  let result;try{result=submitWordTableOperation(this.database,workspace,id,currentUser.id,input as WordTableOperation);}catch(error){if(error instanceof WordTableOperationError)throw new RequestError(error.message,error.status);throw error;}
  if(!result.repeated){this.push(workspace,id,result.update);this.edited(workspace,id,req);}json({update:Buffer.from(result.update).toString('base64'),repeated:result.repeated,...(result.focus?{focus:result.focus}:{})});return;
 }
 if(route==='/office-sheet-operation'&&req.method==='POST'){
  this.writable(workspace,req);const id=text(url.searchParams.get('id'),200);this.documentAccess(workspace,id,user.id,true);
  const input=this.officeOperation(req,workspace,id,await body(req,256*1024));
  this.writable(workspace,req);const currentUser=this.require(req);this.documentAccess(workspace,id,currentUser.id,true);
  const resource=this.database.db.prepare('SELECT kind,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {kind:string;deleted:number}|undefined;
  if(!resource||resource.deleted||resource.kind!=='spreadsheet')throw new RequestError('表格不存在',404);
  let result;try{result=submitSheetOperation(this.database,workspace,id,currentUser.id,input);}catch(error){if(error instanceof SheetOperationError)throw new RequestError(error.message,error.status);throw error;}
  if(!result.repeated){this.push(workspace,id,result.update);this.edited(workspace,id,req);}
  json({update:Buffer.from(result.update).toString('base64'),repeated:result.repeated,before:result.before,after:result.after});return;
 }
 if(route==='/office-resources'&&req.method==='GET'){
  const includeDeleted=url.searchParams.get('deleted')==='1';
  const rows=this.database.db.prepare(`SELECT id,kind,name,mime,source_format AS sourceFormat,original_asset_id AS originalAssetId,model_version AS modelVersion,generation,created_by AS createdBy,created,updated,deleted FROM cloud_office_resources WHERE workspace=? ${includeDeleted?'':'AND deleted=0'} ORDER BY updated DESC`).all(workspace);
  json(rows.filter(row=>{try{this.documentAccess(workspace,(row as {id:string}).id,user.id);return true;}catch(error){if(error instanceof RequestError&&error.status===404)return false;throw error;}}));return;
 }
 if(route==='/office-resources'&&req.method==='POST'){
  this.writable(workspace,req);const input=JSON.parse((await body(req,32768)).toString());
  const kind=input.kind,allowed=['word','spreadsheet','pdf','presentation','text'];if(!allowed.includes(kind))throw new RequestError('办公文件类型无效');
  const name=text(input.name,240),mime=text(input.mime,200),sourceFormat=text(input.sourceFormat,20).toLowerCase(),asset=input.originalAssetId==null?null:text(input.originalAssetId,200),id='doc:'+randomUUID(),now=Date.now(),generation=randomUUID();
  this.writable(workspace,req);if(asset){const owned=this.database.db.prepare('SELECT 1 FROM assets WHERE scope=? AND workspace=? AND id=?').get('cloud:'+workspace,workspace,asset);const binding=this.database.db.prepare('SELECT doc,user FROM cloud_asset_owners WHERE workspace=? AND id=?').get(workspace,asset) as {doc:string|null;user:string}|undefined;if(!owned||!binding||binding.user!==user.id)throw new RequestError('原始文件不存在或无权访问',404);if(binding.doc)throw new RequestError('原件已归属其他文档，请重新上传',409);}
  if(kind==='pdf'&&!asset){const resource=await importOfficeOriginal(this.database,workspace,user.id,name,createBlankPdf(),(format,input)=>this.officeParser.parse(format,input));json(resource,201);return;}
  const initial=asset?undefined:createOfficeState(kind);
  this.database.db.transaction(()=>{this.database.db.prepare('INSERT INTO cloud_office_resources(workspace,id,kind,name,mime,source_format,original_asset_id,model_version,generation,created_by,created,updated,deleted) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,0)').run(workspace,id,kind,name,mime,sourceFormat,asset,1,generation,user.id,now,now);if(asset)this.database.db.prepare('UPDATE cloud_asset_owners SET doc=? WHERE workspace=? AND id=?').run(id,workspace,asset);if(initial)this.database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run('cloud:'+workspace,id,Buffer.from(initial));}).immediate();
  json({id,kind,name,mime,sourceFormat,originalAssetId:asset,modelVersion:1,generation,createdBy:user.id,created:now,updated:now,deleted:0},201);return;
 }
 if(route==='/office-resource'&&['GET','PATCH','DELETE'].includes(req.method??'')){
  const id=text(url.searchParams.get('id'),200);if(!id.startsWith('doc:'))throw new RequestError('办公资源标识无效');
  const resource=this.database.db.prepare('SELECT id,kind,name,mime,source_format AS sourceFormat,original_asset_id AS originalAssetId,model_version AS modelVersion,generation,created_by AS createdBy,created,updated,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {id:string;kind:string;name:string;mime:string;sourceFormat:string;generation:string;deleted:number}|undefined;
  if(!resource||resource.deleted)throw new RequestError('办公资源不存在或已删除',404);this.documentAccess(workspace,id,user.id,req.method!=='GET');
  if(req.method==='GET'){json(resource);return;}
  if(req.method==='PATCH'){const input=JSON.parse((await body(req,8192)).toString()),name=text(input.name,240),now=Date.now();this.officeGeneration(workspace,id,req.headers['x-mydoc-office-generation']);this.database.db.prepare('UPDATE cloud_office_resources SET name=?,updated=? WHERE workspace=? AND id=?').run(name,now,workspace,id);json({ok:true,name,updated:now});return;}
  this.officeGeneration(workspace,id,req.headers['x-mydoc-office-generation']);this.database.db.transaction(()=>{this.database.db.prepare('UPDATE cloud_office_resources SET deleted=1,updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);this.database.db.prepare('INSERT OR IGNORE INTO cloud_deleted(workspace,id) VALUES (?,?)').run(workspace,id);}).immediate();this.disconnect(workspace);json({ok:true});return;
 }
 if(route==='/document-visibility'&&req.method==='GET'){
  const raw=url.searchParams.getAll('id');if(raw.length>1000)throw new RequestError('文档数量过多');
  const ids=raw.map(id=>text(id,200));if(ids.some(id=>!id.startsWith('doc:')))throw new RequestError('文档标识无效');
  const result=ids.map(id=>{if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,id)||!this.database.load('cloud:'+workspace,id))return {id,allowed:false};try{return {id,allowed:true,role:this.documentAccess(workspace,id,user.id).role};}catch(error){if(error instanceof RequestError&&error.status===404)return {id,allowed:false};throw error;}});
  json(result);return;
 }
 if(route==='/invite'&&req.method==='POST'){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString()),email=text(input.email,254).toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new RequestError('邮箱无效');const token=randomBytes(24).toString('hex');this.authorize(workspace,this.require(req).id,true);const role=input.role==='viewer'?'viewer':'member',expiresAt=Date.now()+7*86400000;
  // 重新生成即替换：同邮箱的旧口令一并删除，避免旧口令继续可用。
  this.database.db.transaction(()=>{this.database.db.prepare('DELETE FROM cloud_invites WHERE workspace=? AND email=?').run(workspace,email);this.database.db.prepare('INSERT INTO cloud_invites VALUES (?,?,?,?,?)').run(hash(token),workspace,email,role,expiresAt);}).immediate();json({token,role,expiresAt});return;}
 if(route==='/member'&&['PATCH','DELETE'].includes(req.method??'')){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString()),target=text(input.user,100),old=this.member(workspace,target);member=this.authorize(workspace,this.require(req).id,true);if(!old||old.role==='owner'||member.role!=='owner'&&old.role==='admin')throw new RequestError('不能修改此成员',403);if(req.method==='DELETE'){this.database.db.prepare('DELETE FROM cloud_members WHERE workspace=? AND user=?').run(workspace,target);this.disconnect(workspace,target);}else{if(member.role!=='owner'||!['admin','member','viewer'].includes(input.role))throw new RequestError('角色无效或权限不足',403);this.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run(input.role,workspace,target);this.disconnect(workspace,target);}json({ok:true});return;}
 if(route==='/document-permissions'&&req.method==='GET'){this.authorize(workspace,this.require(req).id,true);const doc=text(url.searchParams.get('doc'),200);this.permissionDocument(workspace,doc);json(this.database.db.prepare('SELECT p.user,u.name,u.email,p.role FROM cloud_document_permissions p JOIN cloud_members m ON m.workspace=p.workspace AND m.user=p.user JOIN cloud_users u ON u.id=p.user WHERE p.workspace=? AND p.doc=?').all(workspace,doc));return;}
 if(route==='/document-permissions'&&['PUT','DELETE'].includes(req.method??'')){this.authorize(workspace,this.require(req).id,true);const input=JSON.parse((await body(req)).toString()),doc=text(input.doc,200),target=text(input.user,100);this.authorize(workspace,this.require(req).id,true);this.permissionDocument(workspace,doc);const targetMember=this.member(workspace,target);if(!targetMember||['owner','admin'].includes(targetMember.role)||target===user.id)throw new RequestError('文档权限目标无效');if(req.method==='DELETE')this.database.db.prepare('DELETE FROM cloud_document_permissions WHERE workspace=? AND doc=? AND user=?').run(workspace,doc,target);else{if(!['viewer','editor','denied'].includes(input.role)||targetMember.role==='viewer'&&input.role==='editor')throw new RequestError('文档权限不能超出工作空间角色');this.database.db.prepare('INSERT INTO cloud_document_permissions(workspace,doc,user,role) VALUES (?,?,?,?) ON CONFLICT(workspace,doc,user) DO UPDATE SET role=excluded.role').run(workspace,doc,target,input.role);}this.disconnect(workspace,target);json({ok:true});return;}

 if(route==='/leave'&&req.method==='POST'){if(member.role==='owner')throw new RequestError('请先转移所有权',403);this.database.db.prepare('DELETE FROM cloud_members WHERE workspace=? AND user=?').run(workspace,user.id);this.disconnect(workspace,user.id);json({ok:true});return;}
 if(route==='/transfer'&&req.method==='POST'){if(member.role!=='owner')throw new RequestError('需要所有者权限',403);const target=text(JSON.parse((await body(req)).toString()).user,100);if(target===user.id||!this.member(workspace,target))throw new RequestError('请选择其他现有成员');this.database.db.transaction(()=>{if(this.authorize(workspace,user.id).role!=='owner'||!this.member(workspace,target))throw new RequestError('成员或权限已变更',403);this.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run('admin',workspace,user.id);this.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run('owner',workspace,target);}).immediate();json({ok:true});return;}
 // 邀请列表包含角色与过期状态：过期口令已不可接受，但仍展示供管理员确认或重新生成。
 if(route==='/invites'&&req.method==='GET'){this.authorize(workspace,this.require(req).id,true);json(this.database.db.prepare('SELECT token AS id,email,role,expires AS expiresAt FROM cloud_invites WHERE workspace=? ORDER BY expires DESC').all(workspace).map(row=>{const invite=row as {id:string;email:string;role:string;expiresAt:number};return {...invite,expired:invite.expiresAt<=Date.now()};}));return;}
 if(route==='/invite'&&req.method==='DELETE'){this.authorize(workspace,this.require(req).id,true);const id=text(JSON.parse((await body(req)).toString()).id,64);this.authorize(workspace,this.require(req).id,true);this.database.db.prepare('DELETE FROM cloud_invites WHERE workspace=? AND token=?').run(workspace,id);json({ok:true});return;}
 if(route==='/presence'&&req.method==='GET'){
  const doc=text(url.searchParams.get('id'),200);if(!this.canReadPresence(workspace,doc,user.id))throw new RequestError('文档不存在或无权访问',404);
  json([...this.presence.values()].filter(p=>p.workspace===workspace&&p.doc===doc&&p.at>=Date.now()-30000&&this.canReadPresence(workspace,doc,p.user)&&!!this.database.db.prepare('SELECT 1 FROM cloud_sessions WHERE token=? AND expires>?').get(p.session,Date.now())).map(p=>({client:p.client,clock:p.clock,state:p.state})));return;
 }
 if(route==='/presence'&&req.method==='POST'){const doc=text(url.searchParams.get('id'),200),input=JSON.parse((await body(req,8192)).toString());const current=this.require(req);this.documentAccess(workspace,doc,current.id);if(!this.canReadPresence(workspace,doc,current.id))throw new RequestError('文档不存在或无权访问',404);if(!/^doc:/.test(doc)&&doc!=='workspace:'+workspace)throw new RequestError('文档标识无效');if(!Number.isSafeInteger(input.client)||input.client<0||input.client>0xffffffff||!Number.isSafeInteger(input.clock)||input.clock<0||typeof input.active!=='boolean')throw new RequestError('在线状态无效');if(input.cursor!==null&&(!input.cursor||!position(input.cursor.anchor)||!position(input.cursor.head)))throw new RequestError('光标位置无效');if(input.focus!==undefined&&input.focus!==null&&!validOfficeFocus(input.focus,position))throw new RequestError('协作位置无效');const key=workspace+':'+doc+':'+input.client,old=this.presence.get(key);if(old&&old.user!==user.id)throw new RequestError('在线状态标识冲突',409);if(!old&&this.presence.size>=10000)throw new RequestError('在线状态过多，请稍后重试',429);if(old&&input.clock<old.clock){json({ok:true});return;}const colors=['#3159cb','#b54708','#027a48','#7a5af8'];const state=input.active?{user:{name:user.name,color:colors[parseInt(hash(user.id).slice(0,8),16)%colors.length]},cursor:input.cursor,...(input.focus?{focus:input.focus}:{})}:null;const p={workspace,doc,client:input.client,clock:input.clock,user:user.id,session:hash(this.token(req)!),sessionId:current.session,state,at:Date.now()};this.presence.set(key,p);this.broadcastPresence(p);json({ok:true});return;}
 if(route==='/assets'&&req.method==='DELETE'){if(member.role!=='owner')throw new RequestError('需要所有者权限',403);if(this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? LIMIT 1').get(workspace))throw new RequestError('空间包含不可变办公原件，请使用工作空间删除接口',409);this.database.db.transaction(()=>{this.database.attachments.remove('cloud:'+workspace);this.database.db.prepare('DELETE FROM cloud_asset_owners WHERE workspace=?').run(workspace);}).immediate();json({ok:true});return;}
 if(route==='/asset'){
  const id=text(url.searchParams.get('id'),200),scope='cloud:'+workspace,document=url.searchParams.has('doc')?text(url.searchParams.get('doc'),200):null;
  if(document&&(!document.startsWith('doc:')||document.length<=4))throw new RequestError('文档标识无效');
  const check=()=>{
   const current=this.require(req),role=this.authorize(workspace,current.id),write=req.method!=='GET';
   if(write){this.writable(workspace,req);if(document)this.officeGeneration(workspace,document,req.headers['x-mydoc-office-generation']);}
   const binding=this.database.db.prepare('SELECT doc,user FROM cloud_asset_owners WHERE workspace=? AND id=?').get(workspace,id) as {doc:string|null;user:string}|undefined;
   if(binding?.doc){
    if(document&&document!==binding.doc)throw new RequestError('附件不存在或无权访问',404);
    if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,binding.doc))throw new RequestError('附件不存在或无权访问',404);
    this.documentAccess(workspace,binding.doc,current.id,write);
   }else if(binding){
    if(binding.user!==current.id)throw new RequestError('附件不存在或无权访问',404);
    if(document)throw new RequestError('跨文档使用附件需要复制为新附件',409);
   }else{
    const exists=this.database.db.prepare('SELECT 1 FROM assets WHERE scope=? AND workspace=? AND id=?').get(scope,workspace,id);
    if(exists&&!['owner','admin'].includes(role.role))throw new RequestError('旧附件尚未绑定文档，请联系空间管理员',404);
    if(document){if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,document))throw new RequestError('文档已永久删除',404);this.documentAccess(workspace,document,current.id,write);}
   }
   if(write&&binding?.doc&&this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,binding.doc))throw new RequestError('办公图片不可覆盖或删除，请插入新图片',409);
   if(write&&this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND original_asset_id=?').get(workspace,id))throw new RequestError('办公原件不可覆盖或删除，请创建新资源',409);
   return current;
  };
  check();
  if(req.method==='GET'){const a=this.database.attachments.get(scope,workspace,id) as {name:string;mime:string;data:Buffer}|undefined;if(!a){json({error:'附件不存在'},404);return;}json({id,name:a.name,mime:a.mime,size:a.data.length,data:Buffer.from(a.data).toString('base64')});return;}
  if(req.method==='PUT'){const data=await body(req,20*1024*1024);this.database.db.transaction(()=>{const current=check();if(document&&this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,document)){if(officeImageMime(data)!==text(url.searchParams.get('mime'),200))throw new RequestError('图片内容与类型不匹配');}this.database.attachments.put(scope,workspace,id,text(url.searchParams.get('name'),1000),text(url.searchParams.get('mime'),200),data);this.database.db.prepare('INSERT OR IGNORE INTO cloud_asset_owners VALUES (?,?,?,?)').run(workspace,id,document,current.id);}).immediate();json({ok:true});return;}
  if(req.method==='DELETE'){this.database.db.transaction(()=>{check();this.database.attachments.remove(scope,workspace,id);this.database.db.prepare('DELETE FROM cloud_asset_owners WHERE workspace=? AND id=?').run(workspace,id);}).immediate();json({ok:true});return;}
 }
 const id=text(url.searchParams.get('id'),200);if(!(route==='/events'&&id==='*')&&!/^(workspace:|doc:)/.test(id))throw new RequestError('文档标识无效');const scope='cloud:'+workspace;
 const officeResource=this.database.db.prepare('SELECT kind,source_format AS sourceFormat,deleted FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {kind:string;sourceFormat:string;deleted:number}|undefined;
 if(officeResource?.deleted)throw new RequestError('办公资源已删除，请刷新工作空间',410);
 if(officeResource&&officeResource.kind!=='text'&&!(officeResource.kind==='pdf'&&this.database.load(scope,id)||officeResource.kind==='presentation'&&this.database.load(scope,id)||officeResource.kind==='word'&&this.database.load(scope,id)||officeResource.kind==='spreadsheet'&&['csv','tsv','xlsx','xlsm','xltx','xltm'].includes(officeResource.sourceFormat)&&this.database.load(scope,id)))throw new RequestError('该办公文件类型暂未接入协同编辑',415);
 if(/^doc:/.test(id)){if(route==='/document'&&req.method==='POST')this.documentAccess(workspace,id,user.id,true);else if(['GET','POST'].includes(req.method??'')&&['/document','/sync','/events'].includes(route))this.documentAccess(workspace,id,user.id);}
 if(route==='/sync'&&req.method==='POST'){const vector=await body(req,65536),generation=this.officeGeneration(workspace,id,req.headers['x-mydoc-office-generation']);this.require(req);const current=this.documentAccess(workspace,id,user.id);if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,id))throw new RequestError('文档已永久删除，请导出草稿',410);const doc=new Y.Doc();try{const bytes=this.database.load(scope,id);if(bytes)Y.applyUpdate(doc,bytes);let update:Uint8Array;try{update=Y.encodeStateAsUpdate(doc,vector);}catch{throw new RequestError('同步向量无效');}json({role:current.role,generation,update:Buffer.from(update).toString('base64'),vector:Buffer.from(Y.encodeStateVector(doc)).toString('base64')});}finally{doc.destroy();}return;}
 if(route==='/document'&&req.method==='GET'){const bytes=this.database.load(scope,id),generation=(this.database.db.prepare('SELECT generation FROM cloud_office_resources WHERE workspace=? AND id=?').get(workspace,id) as {generation:string}|undefined)?.generation;res.writeHead(bytes?200:204,{'Content-Type':'application/octet-stream',...(generation?{'X-Mydoc-Office-Generation':generation}:{})});res.end(bytes??undefined);return;}
 if(route==='/document'&&req.method==='POST'){if(officeResource?.kind==='pdf'){const previous=new Y.Doc();try{const state=this.database.load(scope,id);if(state)Y.applyUpdate(previous,state);if(previous.share.has('pdfOperations'))throw new RequestError('PDF 批注请使用专用操作接口',409);}finally{previous.destroy();}}if(officeResource?.kind==='spreadsheet')throw new RequestError('表格请使用单元格或结构操作接口提交',409);this.writable(workspace,req);const packet=await body(req,16*1024*1024+256),unpacked=officeResource?unpackOfficeUpdate(packet):{update:packet,generation:undefined},bytes=unpacked.update;if(bytes.length>16*1024*1024)throw new RequestError('更新超过 16 MB');if(officeResource)this.officeGeneration(workspace,id,unpacked.generation??req.headers['x-mydoc-office-generation']);this.writable(workspace,req);this.documentAccess(workspace,id,this.require(req).id,true);if(this.database.db.prepare('SELECT 1 FROM cloud_deleted WHERE workspace=? AND id=?').get(workspace,id))throw new RequestError('文档已永久删除，请刷新工作空间',410);// Clients re-send already-known updates when reconciling; only updates that change the file count as edits.
 let changed=false;if(officeResource){try{assertOfficeCommentsUnchanged(this.database.load(scope,id),bytes);if(officeResource.kind==='text')validateOfficeTextUpdate(this.database.load(scope,id),bytes);else if(officeResource.kind==='pdf')validatePdfUpdate(this.database.load(scope,id),bytes);else if(officeResource.kind==='presentation')validatePresentationUpdate(this.database.load(scope,id),bytes);else if(officeResource.kind==='word')validateWordUpdate(this.database.load(scope,id),bytes,{lockTables:!this.localContexts.has(req),...(this.localContexts.has(req)?{}:{actor:this.require(req).id})});else validateSpreadsheetUpdate(this.database.load(scope,id),bytes);const model=new Y.Doc();try{const previous=this.database.load(scope,id);if(previous)Y.applyUpdate(model,previous);model.on('update',()=>{changed=true;});Y.applyUpdate(model,bytes);readResourceMedia(this.database,workspace,id,officeResource.kind as import('../../../../common/nbstore/src/office').OfficeKind,Y.encodeStateAsUpdate(model));}finally{model.destroy();}}catch(error){throw new RequestError(error instanceof Error?error.message:'办公文件更新无效');}}this.database.db.transaction(()=>{this.database.append(scope,id,bytes);if(officeResource)this.database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);}).immediate();this.push(workspace,id,bytes);if(officeResource&&changed)this.edited(workspace,id,req);json({ok:true});return;}
 if(route==='/document'&&req.method==='DELETE'){if(officeResource)throw new RequestError('请使用办公资源删除接口',409);this.writable(workspace,req);this.documentAccess(workspace,id,user.id,true);if(id.startsWith('workspace:'))throw new RequestError('请使用工作空间删除接口',403);this.database.db.transaction(()=>{this.database.db.prepare('INSERT OR IGNORE INTO cloud_deleted VALUES (?,?)').run(workspace,id);this.database.db.prepare('DELETE FROM documents WHERE scope=? AND id=?').run(scope,id);this.database.db.prepare('DELETE FROM cloud_document_permissions WHERE workspace=? AND doc=?').run(workspace,id);}).immediate();this.disconnect(workspace);json({ok:true});return;}
 if(route==='/events'&&req.method==='GET'){if([...this.streams.values()].filter(s=>s.user===user.id).length>=30)throw new RequestError('同步连接过多',429);res.writeHead(200,{'Content-Type':'text/event-stream','Connection':'keep-alive','X-Accel-Buffering':'no'});this.streams.set(res,{user:user.id,workspace,doc:id,expires:user.expires,session:hash(this.token(req)!),sessionId:user.session});for(const p of this.presence.values())if(p.workspace===workspace&&(id==='*'||id===p.doc)&&this.canReadPresence(workspace,p.doc,user.id)&&this.canReadPresence(workspace,p.doc,p.user)&&p.at>=Date.now()-30000)res.write(`data: ${JSON.stringify({id:p.doc,presence:{client:p.client,clock:p.clock,state:p.state}})}\n\n`);res.write(': connected\n\n');req.on('close',()=>this.streams.delete(res));return;}
 json({error:'接口不存在'},404);
 }catch(error){if(res.headersSent){res.end();return;}json({error:error instanceof RequestError?error.message:'请求失败，请稍后重试',...(error instanceof RequestError&&error.code?{code:error.code}:{})},error instanceof RequestError?error.status:500);}};
 close(){this.officeExports.close();this.officeParser.close();clearInterval(this.timer);for(const res of this.streams.keys())res.end();this.streams.clear();if(this.ownsDatabase)this.database.close();}
}
