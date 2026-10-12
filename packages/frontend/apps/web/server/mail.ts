import {connect as netConnect,type Socket} from 'node:net';
import {connect as tlsConnect} from 'node:tls';
import {hostname} from 'node:os';

/**
 * 账户邮件发送：邮箱验证、找回密码与更改邮箱都走这里。
 * 未配置 SMTP 时不做任何静默降级，status() 明确返回未配置原因，
 * 调用方据此改为把一次性口令交给运维或当前已登录用户，不假装邮件已发出。
 */
export type MailMessage={to:string;subject:string;text:string};
export interface MailTransport{send(message:MailMessage):Promise<void>}
export type MailConfig={host:string;port:number;secure:boolean;user?:string;pass?:string;from:string};
export type MailStatus={configured:boolean;from?:string;host?:string;reason?:string};

export function mailConfigFromEnv(env:Record<string,string|undefined>=process.env):{config?:MailConfig;reason?:string}{
 const host=env.MYDOC_SMTP_HOST?.trim(),from=env.MYDOC_MAIL_FROM?.trim();
 if(!host&&!from)return {reason:'未设置 MYDOC_SMTP_HOST 与 MYDOC_MAIL_FROM'};
 if(!host)return {reason:'未设置 MYDOC_SMTP_HOST'};
 if(!from)return {reason:'未设置 MYDOC_MAIL_FROM'};
 const secure=env.MYDOC_SMTP_SECURE==='1'||env.MYDOC_SMTP_SECURE==='true';
 const port=Number(env.MYDOC_SMTP_PORT??(secure?465:587));
 if(!Number.isInteger(port)||port<1||port>65535)return {reason:'MYDOC_SMTP_PORT 不是有效端口'};
 const user=env.MYDOC_SMTP_USER?.trim()||undefined,pass=env.MYDOC_SMTP_PASS||undefined;
 if(user&&!pass)return {reason:'设置了 MYDOC_SMTP_USER 但缺少 MYDOC_SMTP_PASS'};
 return {config:{host,port,secure,user,pass,from}};
}

const word=(value:string)=>/^[\x20-\x7e]*$/.test(value)?value:'=?UTF-8?B?'+Buffer.from(value,'utf8').toString('base64')+'?=';

/** 极简 SMTP 客户端：EHLO、可选 STARTTLS、AUTH LOGIN、单收件人投递。不做队列与重试。 */
export class SmtpTransport implements MailTransport {
 constructor(private config:MailConfig,private timeout=15000){}
 async send(message:MailMessage){
  const {config}=this;let socket:Socket=await this.open(config.secure);
  const reader=createReader();let detach=reader.attach(socket);
  const say=async(line:string,expect:number[])=>{socket.write(line+'\r\n');return reader.expect(expect,this.timeout);};
  try{
   await reader.expect([220],this.timeout);
   let greeting=await say('EHLO '+(hostname()||'localhost'),[250]);
   if(!config.secure&&/\bSTARTTLS\b/i.test(greeting.text)){
    await say('STARTTLS',[220]);detach();
    socket=await new Promise<Socket>((resolve,reject)=>{const upgraded=tlsConnect({socket,servername:config.host},()=>resolve(upgraded as unknown as Socket));upgraded.once('error',reject);});
    detach=reader.attach(socket);greeting=await say('EHLO '+(hostname()||'localhost'),[250]);
   }
   if(config.user&&config.pass){
    if(!config.secure&&!(socket as unknown as {encrypted?:boolean}).encrypted)throw new Error('SMTP 服务器不支持 STARTTLS，拒绝在明文连接上发送凭据');
    await say('AUTH LOGIN',[334]);await say(Buffer.from(config.user,'utf8').toString('base64'),[334]);await say(Buffer.from(config.pass,'utf8').toString('base64'),[235]);
   }
   await say('MAIL FROM:<'+config.from+'>',[250]);
   await say('RCPT TO:<'+message.to+'>',[250,251]);
   await say('DATA',[354]);
   const body=Buffer.from(message.text,'utf8').toString('base64').replace(/(.{76})/g,'$1\r\n');
   const headers=['From: '+config.from,'To: '+message.to,'Subject: '+word(message.subject),'Date: '+new Date().toUTCString(),'MIME-Version: 1.0','Content-Type: text/plain; charset=utf-8','Content-Transfer-Encoding: base64'];
   socket.write(headers.join('\r\n')+'\r\n\r\n'+body+'\r\n.\r\n');
   await reader.expect([250],this.timeout);
   socket.write('QUIT\r\n');
  }finally{detach();socket.destroy();}
 }
 private open(secure:boolean){return new Promise<Socket>((resolve,reject)=>{
  const socket=secure?tlsConnect({host:this.config.host,port:this.config.port,servername:this.config.host}) as unknown as Socket:netConnect({host:this.config.host,port:this.config.port});
  socket.setTimeout(this.timeout,()=>socket.destroy(new Error('SMTP 连接超时')));
  socket.once(secure?'secureConnect':'connect',()=>resolve(socket));socket.once('error',reject);
 });}
}

/** 按 SMTP 回复格式切分：多行回复以“代码+空格”结尾，读满一整条才算一次应答。 */
function createReader(){
 let buffer='',waiting:{resolve:(value:{code:number;text:string})=>void;reject:(error:Error)=>void;expect:number[];timer:ReturnType<typeof setTimeout>}|undefined;
 const flush=()=>{
  if(!waiting)return;
  const match=buffer.match(/^(?:\d{3}-[^\n]*\n)*(\d{3}) [^\n]*\n/);
  if(!match)return;
  const text=buffer.slice(0,match[0].length);buffer=buffer.slice(match[0].length);
  const code=Number(match[1]),pending=waiting;waiting=undefined;clearTimeout(pending.timer);
  if(pending.expect.includes(code))pending.resolve({code,text});else pending.reject(new Error('SMTP 服务器返回 '+text.trim()));
 };
 const fail=(error:Error)=>{if(waiting){clearTimeout(waiting.timer);waiting.reject(error);waiting=undefined;}};
 return {
  attach(socket:Socket){
   const data=(chunk:Buffer|string)=>{buffer+=chunk.toString();flush();},error=(cause:Error)=>fail(cause),close=()=>fail(new Error('SMTP 连接已关闭'));
   socket.on('data',data);socket.on('error',error);socket.on('close',close);
   return ()=>{socket.off('data',data);socket.off('error',error);socket.off('close',close);};
  },
  expect(expect:number[],timeout:number){return new Promise<{code:number;text:string}>((resolve,reject)=>{
   waiting={resolve,reject,expect,timer:setTimeout(()=>fail(new Error('等待 SMTP 应答超时')),timeout)};flush();
  });},
 };
}

export class MailService {
 private transport?:MailTransport;readonly status:MailStatus;
 constructor(options:{config?:MailConfig;reason?:string;transport?:MailTransport}={}){
  const resolved=options.transport?{config:options.config,reason:undefined}:(options.config?{config:options.config}:mailConfigFromEnv());
  if(options.transport){this.transport=options.transport;this.status={configured:true,from:options.config?.from??'mydoc@localhost',host:options.config?.host??'注入的发送通道'};}
  else if(resolved.config){this.transport=new SmtpTransport(resolved.config);this.status={configured:true,from:resolved.config.from,host:resolved.config.host+':'+resolved.config.port};}
  else this.status={configured:false,reason:options.reason??resolved.reason??'邮件服务未配置'};
 }
 async send(message:MailMessage){
  if(!this.transport)throw new Error('邮件服务未配置：'+(this.status.reason??'缺少 SMTP 设置'));
  await this.transport.send(message);
 }
}
