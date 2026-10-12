import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {createServer as createSocketServer} from 'node:net';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {MailService,SmtpTransport,mailConfigFromEnv,type MailMessage} from '../packages/frontend/apps/web/server/mail';

const png=Buffer.concat([Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a]),Buffer.alloc(64,7)]);

async function fixture(mail?:MailService){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-cloud-account-'));
 const cloud=new CloudServer(join(dir,'cloud.sqlite'),{mail:mail??new MailService({reason:'测试未配置邮件服务'})});
 const server=createServer(cloud.handler);
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',input?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(input?{'Content-Type':'application/json'}:{})},body:input?JSON.stringify(input):undefined});
 const register=async(email:string,name:string,agent='Vitest/1.0')=>{const response=await fetch(root+'/register',{method:'POST',headers:{'Content-Type':'application/json','User-Agent':agent},body:JSON.stringify({email,name,password:'account-password-1'})});return {cookie:response.headers.get('set-cookie')!.split(';')[0],user:(await response.json()).user as {id:string;email:string;name:string;verified:boolean}};};
 const login=async(email:string,password:string,agent='Vitest/2.0')=>{const response=await fetch(root+'/login',{method:'POST',headers:{'Content-Type':'application/json','User-Agent':agent},body:JSON.stringify({email,password})});return {status:response.status,cookie:response.headers.get('set-cookie')?.split(';')[0]};};
 const upload=(data:Buffer,type:string,cookie:string)=>fetch(root+'/avatar',{method:'PUT',headers:{Cookie:cookie,'Content-Type':type},body:data});
 return {root,cloud,call,register,login,upload,async close(){await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}};
}

it('显示名称与头像可自助维护，头像限制格式、大小并按魔数校验',async()=>{
 const f=await fixture();
 try{
  const a=await f.register('avatar@example.com','原名');
  expect((await (await f.call('/profile','PATCH',{name:'新名字'},a.cookie)).json()).name).toBe('新名字');
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.name).toBe('新名字');
  // 空名称被拒绝，不写入数据库。
  expect((await f.call('/profile','PATCH',{name:'   '},a.cookie)).status).toBe(400);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.name).toBe('新名字');

  expect((await f.call('/session','GET',undefined,a.cookie)).ok).toBe(true);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.hasAvatar).toBe(false);
  const stored=await f.upload(png,'image/png',a.cookie);
  expect(stored.status).toBe(200);
  expect((await stored.json()).bytes).toBe(png.length);
  const fetched=await f.call('/avatar','GET',undefined,a.cookie);
  expect(fetched.headers.get('content-type')).toBe('image/png');
  expect(Buffer.from(await fetched.arrayBuffer()).equals(png)).toBe(true);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.hasAvatar).toBe(true);

  // 声明 PNG 但内容不是图片：按魔数拒绝。
  expect((await f.upload(Buffer.from('not an image at all'),'image/png',a.cookie)).status).toBe(400);
  // 不支持的类型与超限大小分别拒绝。
  expect((await f.upload(png,'image/gif',a.cookie)).status).toBe(400);
  expect((await f.upload(Buffer.concat([png,Buffer.alloc(300*1024)]),'image/png',a.cookie)).status).toBe(413);
  // 原头像仍在，失败的上传没有破坏已有数据。
  expect(Buffer.from(await (await f.call('/avatar','GET',undefined,a.cookie)).arrayBuffer()).equals(png)).toBe(true);

  expect((await f.call('/avatar','DELETE',{},a.cookie)).status).toBe(200);
  expect((await f.call('/avatar','GET',undefined,a.cookie)).status).toBe(204);
  // 未登录不能读写头像。
  expect((await f.call('/avatar','GET')).status).toBe(401);
 }finally{await f.close();}
});

it('改密校验当前密码、让其他设备重新登录，当前会话继续可用',async()=>{
 const f=await fixture();
 try{
  const a=await f.register('password@example.com','甲','Desktop/1.0');
  const other=await f.login('password@example.com','account-password-1','Phone/1.0');
  expect(other.status).toBe(200);
  expect(((await (await f.call('/sessions','GET',undefined,a.cookie)).json()) as unknown[]).length).toBe(2);

  // 当前密码不对：拒绝且不改动密码。
  expect((await f.call('/password','POST',{current:'wrong-password',password:'next-password-1'},a.cookie)).status).toBe(401);
  expect((await f.login('password@example.com','account-password-1')).status).toBe(200);
  // 新密码过短、与原密码相同分别拒绝。
  expect((await f.call('/password','POST',{current:'account-password-1',password:'short'},a.cookie)).status).toBe(400);
  expect((await f.call('/password','POST',{current:'account-password-1',password:'account-password-1'},a.cookie)).status).toBe(400);

  const changed=await f.call('/password','POST',{current:'account-password-1',password:'next-password-1'},a.cookie);
  expect(changed.status).toBe(200);
  expect((await changed.json()).signedOutOthers).toBe(true);
  const rotated=changed.headers.get('set-cookie')!.split(';')[0];
  // 当前会话换了新 Cookie：新 Cookie 可用，旧 Cookie 失效。
  expect((await f.call('/session','GET',undefined,rotated)).ok).toBe(true);
  expect((await (await f.call('/session','GET',undefined,rotated)).json()).user.email).toBe('password@example.com');
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user).toBe(null);
  // 其他设备会话被清理。
  expect((await (await f.call('/session','GET',undefined,other.cookie!)).json()).user).toBe(null);
  expect(((await (await f.call('/sessions','GET',undefined,rotated)).json()) as unknown[]).length).toBe(1);
  // 旧密码不能再登录，新密码可以。
  expect((await f.login('password@example.com','account-password-1')).status).toBe(401);
  expect((await f.login('password@example.com','next-password-1')).status).toBe(200);
 }finally{await f.close();}
});

it('登录设备列表记录设备信息，可注销指定会话或其他全部会话',async()=>{
 const f=await fixture();
 try{
  const a=await f.register('sessions@example.com','甲','Desktop/1.0');
  const phone=await f.login('sessions@example.com','account-password-1','Phone/1.0');
  const tablet=await f.login('sessions@example.com','account-password-1','Tablet/1.0');
  const list=await (await f.call('/sessions','GET',undefined,a.cookie)).json() as {id:string;agent:string;current:boolean;createdAt:number;lastSeenAt:number;expiresAt:number}[];
  expect(list.length).toBe(3);
  expect(list.filter(row=>row.current).length).toBe(1);
  expect(list.map(row=>row.agent).sort()).toEqual(['Desktop/1.0','Phone/1.0','Tablet/1.0']);
  // 列表不泄露 token 或其哈希。
  expect(JSON.stringify(list)).not.toContain('token');
  for(const row of list){expect(row.expiresAt).toBeGreaterThan(Date.now());expect(row.createdAt).toBeGreaterThan(0);}

  const phoneSession=list.find(row=>row.agent==='Phone/1.0')!;
  expect((await f.call('/sessions','DELETE',{session:phoneSession.id},a.cookie)).status).toBe(200);
  expect((await (await f.call('/session','GET',undefined,phone.cookie!)).json()).user).toBe(null);
  // 重复注销同一会话：已失效，明确报错而不是假装成功。
  expect((await f.call('/sessions','DELETE',{session:phoneSession.id},a.cookie)).status).toBe(400);
  // 当前会话必须走退出登录。
  expect((await f.call('/sessions','DELETE',{session:list.find(row=>row.current)!.id},a.cookie)).status).toBe(400);
  // 其他账户的会话 ID 不能被注销。
  const b=await f.register('other-sessions@example.com','乙');
  const bList=await (await f.call('/sessions','GET',undefined,b.cookie)).json() as {id:string}[];
  expect((await f.call('/sessions','DELETE',{session:bList[0].id},a.cookie)).status).toBe(400);
  expect((await (await f.call('/session','GET',undefined,b.cookie)).json()).user.email).toBe('other-sessions@example.com');

  const all=await f.call('/sessions','DELETE',{all:true},a.cookie);
  expect((await all.json()).removed).toBe(1);
  expect((await (await f.call('/session','GET',undefined,tablet.cookie!)).json()).user).toBe(null);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.email).toBe('sessions@example.com');
 }finally{await f.close();}
});

it('邮箱验证与更改邮箱使用一次性口令，过期与重复使用均被拒绝',async()=>{
 const f=await fixture();
 try{
  const a=await f.register('verify@example.com','甲');
  expect(a.user.verified).toBe(false);
  // 未配置邮件服务：明确告知未发送，并把口令回传给已登录本人。
  const requested=await (await f.call('/verify','POST',{},a.cookie)).json() as {sent:boolean;token:string;notice:string};
  expect(requested.sent).toBe(false);
  expect(requested.notice).toContain('邮件未发送');
  expect(requested.token).toMatch(/^[0-9a-f]{48}$/);
  // 重发会替换旧口令：旧口令立即失效。
  const resent=await (await f.call('/verify','POST',{},a.cookie)).json() as {token:string};
  expect(resent.token).not.toBe(requested.token);
  expect((await f.call('/verify','PATCH',{token:requested.token},a.cookie)).status).toBe(400);
  expect((await f.call('/verify','PATCH',{token:resent.token},a.cookie)).status).toBe(200);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.verified).toBe(true);
  // 已验证后不再重复发送。
  expect((await f.call('/verify','POST',{},a.cookie)).status).toBe(400);

  const b=await f.register('occupied@example.com','乙');
  // 更改邮箱要校验当前密码、唯一性与是否同邮箱。
  expect((await f.call('/email','POST',{email:'moved@example.com',password:'wrong'},a.cookie)).status).toBe(401);
  expect((await f.call('/email','POST',{email:'occupied@example.com',password:'account-password-1'},a.cookie)).status).toBe(400);
  expect((await f.call('/email','POST',{email:'verify@example.com',password:'account-password-1'},a.cookie)).status).toBe(400);
  const change=await (await f.call('/email','POST',{email:'moved@example.com',password:'account-password-1'},a.cookie)).json() as {token:string;email:string;pendingInvites:number};
  expect(change.email).toBe('moved@example.com');
  expect(change.pendingInvites).toBe(0);
  // 口令只能用一次。
  expect((await (await f.call('/email','PATCH',{token:change.token},a.cookie)).json()).email).toBe('moved@example.com');
  expect((await f.call('/email','PATCH',{token:change.token},a.cookie)).status).toBe(400);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.email).toBe('moved@example.com');
  // 换邮箱后用新地址登录，旧地址不再有效。
  expect((await f.login('moved@example.com','account-password-1')).status).toBe(200);
  expect((await f.login('verify@example.com','account-password-1')).status).toBe(401);
  // 另一账户的口令不能用在当前账户上。
  const bChange=await (await f.call('/email','POST',{email:'b-moved@example.com',password:'account-password-1'},b.cookie)).json() as {token:string};
  expect((await f.call('/email','PATCH',{token:bChange.token},a.cookie)).status).toBe(400);

  // 过期口令：直接改写数据库中的到期时间，模拟 1 小时后再使用。
  const expiring=await (await f.call('/email','POST',{email:'later@example.com',password:'account-password-1'},a.cookie)).json() as {token:string};
  f.cloud.database.db.prepare('UPDATE cloud_tokens SET expires=? WHERE kind=?').run(Date.now()-1000,'email');
  expect((await f.call('/email','PATCH',{token:expiring.token},a.cookie)).status).toBe(400);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user.email).toBe('moved@example.com');
 }finally{await f.close();}
});

it('更改邮箱后发往旧地址的未接受邀请不再可用，需重新邀请',async()=>{
 const f=await fixture();
 try{
  const owner=await f.register('owner-invite@example.com','所有者');
  const guest=await f.register('guest-old@example.com','受邀人');
  const space=await (await f.call('/workspaces','POST',{name:'邀请与邮箱'},owner.cookie)).json() as {id:string};
  const invite=await (await f.call('/invite?workspace='+space.id,'POST',{email:'guest-old@example.com'},owner.cookie)).json() as {token:string};

  // 换邮箱前提示旧地址仍有未接受邀请。
  const change=await (await f.call('/email','POST',{email:'guest-new@example.com',password:'account-password-1'},guest.cookie)).json() as {token:string;pendingInvites:number};
  expect(change.pendingInvites).toBe(1);
  const applied=await (await f.call('/email','PATCH',{token:change.token},guest.cookie)).json() as {staleInvites:number};
  expect(applied.staleInvites).toBe(1);
  // 旧口令绑定旧邮箱，换邮箱后不能接受。
  expect((await f.call('/accept','POST',{token:invite.token},guest.cookie)).status).toBe(400);
  // 对新邮箱重新邀请后可以加入。
  const reissued=await (await f.call('/invite?workspace='+space.id,'POST',{email:'guest-new@example.com'},owner.cookie)).json() as {token:string};
  expect((await f.call('/accept','POST',{token:reissued.token},guest.cookie)).status).toBe(200);
 }finally{await f.close();}
});

it('找回密码不透露邮箱是否存在，重置后清空全部会话且口令不可重复使用',async()=>{
 const sent:MailMessage[]=[];
 const f=await fixture(new MailService({transport:{async send(message){sent.push(message);}}}));
 try{
  const a=await f.register('reset@example.com','甲','Desktop/1.0');
  const phone=await f.login('reset@example.com','account-password-1','Phone/1.0');

  // 不存在的邮箱与存在的邮箱返回相同结果。
  const missing=await f.call('/forgot','POST',{email:'nobody@example.com'});
  const missingBody=await missing.json();
  const present=await f.call('/forgot','POST',{email:'reset@example.com'});
  const presentBody=await present.json();
  expect(missing.status).toBe(present.status);
  expect(missingBody).toEqual(presentBody);
  // 只给真实存在的账户发了邮件，响应体不含口令。
  expect(sent.length).toBe(1);
  expect(sent[0].to).toBe('reset@example.com');
  const token=sent[0].text.match(/[0-9a-f]{48}/)![0];
  expect(JSON.stringify(presentBody)).not.toContain(token);

  expect((await f.call('/reset','POST',{token,password:'short'})).status).toBe(400);
  expect((await f.call('/reset','POST',{token:'0'.repeat(48),password:'reset-password-1'})).status).toBe(400);
  expect((await f.call('/reset','POST',{token,password:'reset-password-1'})).status).toBe(200);
  // 口令一次性。
  expect((await f.call('/reset','POST',{token,password:'another-password-1'})).status).toBe(400);
  // 重置后所有设备都要重新登录。
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user).toBe(null);
  expect((await (await f.call('/session','GET',undefined,phone.cookie!)).json()).user).toBe(null);
  expect((await f.login('reset@example.com','account-password-1')).status).toBe(401);
  expect((await f.login('reset@example.com','reset-password-1')).status).toBe(200);
 }finally{await f.close();}
});

it('账户切换与会话撤销不会暴露上一账户的数据',async()=>{
 const f=await fixture();
 try{
  const a=await f.register('first@example.com','甲');
  const b=await f.register('second@example.com','乙');
  const space=await (await f.call('/workspaces','POST',{name:'甲的空间'},a.cookie)).json() as {id:string};
  await f.upload(png,'image/png',a.cookie);

  // 乙的会话看不到甲的工作空间与概览。
  expect(await (await f.call('/workspaces','GET',undefined,b.cookie)).json()).toEqual([]);
  expect((await f.call('/overview?workspace='+space.id,'GET',undefined,b.cookie)).status).toBe(403);
  expect((await (await f.call('/session','GET',undefined,b.cookie)).json()).user.email).toBe('second@example.com');
  // 账户头显式不符时拒绝，避免切号后旧页面继续读写。
  expect((await fetch(f.root+'/workspaces',{headers:{Cookie:b.cookie,'X-Mydoc-Account':a.user.id}})).status).toBe(409);

  // 注销甲的会话后，该 Cookie 读不到任何账户数据。
  await f.call('/logout','POST',{},a.cookie);
  expect((await (await f.call('/session','GET',undefined,a.cookie)).json()).user).toBe(null);
  expect((await f.call('/workspaces','GET',undefined,a.cookie)).status).toBe(401);
  expect((await f.call('/avatar','GET',undefined,a.cookie)).status).toBe(401);
  // 重新登录得到新会话，数据仍属于本人。
  const again=await f.login('first@example.com','account-password-1');
  expect(((await (await f.call('/workspaces','GET',undefined,again.cookie!)).json()) as {id:string}[])[0].id).toBe(space.id);
 }finally{await f.close();}
});

it('邮件服务未配置时明确报告原因，配置来源于环境变量',async()=>{
 expect(mailConfigFromEnv({}).reason).toContain('MYDOC_SMTP_HOST');
 expect(mailConfigFromEnv({MYDOC_SMTP_HOST:'smtp.example.com'}).reason).toContain('MYDOC_MAIL_FROM');
 expect(mailConfigFromEnv({MYDOC_MAIL_FROM:'mydoc@example.com'}).reason).toContain('MYDOC_SMTP_HOST');
 expect(mailConfigFromEnv({MYDOC_SMTP_HOST:'smtp.example.com',MYDOC_MAIL_FROM:'a@b.com',MYDOC_SMTP_PORT:'abc'}).reason).toContain('端口');
 expect(mailConfigFromEnv({MYDOC_SMTP_HOST:'smtp.example.com',MYDOC_MAIL_FROM:'a@b.com',MYDOC_SMTP_USER:'u'}).reason).toContain('MYDOC_SMTP_PASS');
 const ok=mailConfigFromEnv({MYDOC_SMTP_HOST:'smtp.example.com',MYDOC_MAIL_FROM:'mydoc@example.com'});
 expect(ok.config).toEqual({host:'smtp.example.com',port:587,secure:false,user:undefined,pass:undefined,from:'mydoc@example.com'});
 expect(mailConfigFromEnv({MYDOC_SMTP_HOST:'h',MYDOC_MAIL_FROM:'a@b.com',MYDOC_SMTP_SECURE:'1'}).config?.port).toBe(465);

 const unconfigured=new MailService({reason:'未设置 MYDOC_SMTP_HOST'});
 expect(unconfigured.status.configured).toBe(false);
 await expect(unconfigured.send({to:'a@b.com',subject:'x',text:'y'})).rejects.toThrow('邮件服务未配置');

 // 未配置时服务端如实上报，不假装已发送。
 const f=await fixture(unconfigured);
 try{
  const session=await (await f.call('/session')).json() as {mail:{configured:boolean;reason:string}};
  expect(session.mail.configured).toBe(false);
  expect(session.mail.reason).toContain('MYDOC_SMTP_HOST');
 }finally{await f.close();}
});

it('SMTP 发送按协议完成投递，收件人、主题与正文可被本地假服务器收到',async()=>{
 const received:string[]=[];
 const server=createSocketServer(socket=>{
  let buffer='',inData=false;
  socket.write('220 fake ESMTP ready\r\n');
  socket.on('data',chunk=>{
   buffer+=chunk.toString();
   for(;;){
    const index=buffer.indexOf('\r\n');if(index<0)break;
    const line=buffer.slice(0,index);buffer=buffer.slice(index+2);
    received.push(line);
    if(inData){if(line==='.'){inData=false;socket.write('250 queued\r\n');}continue;}
    if(/^EHLO/i.test(line))socket.write('250-fake greets you\r\n250 SIZE 1048576\r\n');
    else if(/^MAIL FROM/i.test(line)||/^RCPT TO/i.test(line))socket.write('250 ok\r\n');
    else if(/^DATA/i.test(line)){inData=true;socket.write('354 go ahead\r\n');}
    else if(/^QUIT/i.test(line)){socket.write('221 bye\r\n');socket.end();}
    else socket.write('250 ok\r\n');
   }
  });
  socket.on('error',()=>{});
 });
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const port=(server.address() as {port:number}).port;
 try{
  const transport=new SmtpTransport({host:'127.0.0.1',port,secure:false,from:'mydoc@example.com'},5000);
  await transport.send({to:'someone@example.com','subject':'MyDoc 验证口令',text:'一次性口令：abc123'});
  expect(received.some(line=>line.startsWith('MAIL FROM:<mydoc@example.com>'))).toBe(true);
  expect(received.some(line=>line.startsWith('RCPT TO:<someone@example.com>'))).toBe(true);
  const body=received.join('\n');
  expect(body).toContain('Content-Transfer-Encoding: base64');
  // 非 ASCII 主题按 RFC 2047 编码，正文按 base64 传输。
  expect(body).toContain('Subject: =?UTF-8?B?');
  const encoded=body.split('\n').filter(line=>/^[A-Za-z0-9+/=]+$/.test(line)&&line.length>8).join('');
  expect(Buffer.from(encoded,'base64').toString('utf8')).toContain('一次性口令：abc123');

  // 服务器拒绝收件人时如实抛错，不静默成功。
  const rejecting=createSocketServer(socket=>{socket.write('220 fake\r\n');socket.on('data',chunk=>{const line=chunk.toString();if(/^EHLO/i.test(line))socket.write('250 ok\r\n');else if(/^MAIL FROM/i.test(line))socket.write('550 mailbox unavailable\r\n');});socket.on('error',()=>{});});
  await new Promise<void>(resolve=>rejecting.listen(0,'127.0.0.1',resolve));
  const rejectPort=(rejecting.address() as {port:number}).port;
  await expect(new SmtpTransport({host:'127.0.0.1',port:rejectPort,secure:false,from:'mydoc@example.com'},5000).send({to:'x@example.com',subject:'x',text:'y'})).rejects.toThrow('550');
  await new Promise<void>(resolve=>rejecting.close(()=>resolve()));
 }finally{await new Promise<void>(resolve=>server.close(()=>resolve()));}
});

it('频率限制覆盖找回密码、验证与更改邮箱，并在超限后拒绝',async()=>{
 const f=await fixture();
 try{
  const a=await f.register('limit@example.com','甲');
  // 找回密码每分钟 5 次。
  for(let index=0;index<5;index+=1)expect((await f.call('/forgot','POST',{email:'limit@example.com'})).status).toBe(200);
  expect((await f.call('/forgot','POST',{email:'limit@example.com'})).status).toBe(429);
  // 验证邮件每分钟 5 次，与找回密码分别计数。
  for(let index=0;index<5;index+=1)expect((await f.call('/verify','POST',{},a.cookie)).status).toBe(200);
  expect((await f.call('/verify','POST',{},a.cookie)).status).toBe(429);
  // 更改邮箱也独立计数。
  for(let index=0;index<5;index+=1)expect((await f.call('/email','POST',{email:'limit-'+index+'@example.com',password:'account-password-1'},a.cookie)).status).toBe(200);
  expect((await f.call('/email','POST',{email:'limit-x@example.com',password:'account-password-1'},a.cookie)).status).toBe(429);
  // 会话查询不受这些限制影响。
  expect((await f.call('/session','GET',undefined,a.cookie)).status).toBe(200);
 }finally{await f.close();}
});
