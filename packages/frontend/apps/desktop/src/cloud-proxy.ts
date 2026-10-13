import {createHash} from 'node:crypto';
import {request as httpRequest,type IncomingMessage,type ServerResponse} from 'node:http';
import {request as httpsRequest} from 'node:https';
import type {Readable,Writable} from 'node:stream';

export type CloudResponse=Readable&{statusCode?:number;headers:IncomingMessage['headers']};
export type CloudRequestFactory=(url:URL,method:string,headers:Record<string,string>,respond:(response:CloudResponse)=>void)=>{body:Writable;cancel:()=>void};
const nodeCloudRequest:CloudRequestFactory=(url,method,headers,respond)=>{
  const body=(url.protocol==='https:'?httpsRequest:httpRequest)(url,{method,headers},respond);
  return {body,cancel:()=>body.destroy()};
};

export function cloudConnectionError(error:NodeJS.ErrnoException,server:string,platform=process.platform){
  const code=error.code??'UNKNOWN';
  let reason='请检查服务器地址和网络';
  if(platform==='darwin'&&['EACCES','EPERM','EHOSTUNREACH','ERR_ACCESS_DENIED','ERR_NETWORK_ACCESS_DENIED','ERR_ADDRESS_UNREACHABLE'].includes(code))reason='请在系统设置 > 隐私与安全 > 本地网络中允许 MyDoc，并检查网络路由';
  else if(['ECONNREFUSED','ERR_CONNECTION_REFUSED'].includes(code))reason='服务器拒绝连接，请检查服务监听端口';
  else if(['ETIMEDOUT','ESOCKETTIMEDOUT','ERR_TIMED_OUT','ERR_CONNECTION_TIMED_OUT'].includes(code))reason='连接超时，请检查网络和服务器防火墙';
  else if(['ENOTFOUND','EAI_AGAIN','ERR_NAME_NOT_RESOLVED'].includes(code))reason='服务器域名无法解析，请检查地址和 DNS';
  return `无法连接工作空间服务器 ${server}（${code}），${reason}`;
}

export function proxyCloudRequest(req:IncomingMessage,res:ServerResponse,origin:string,requestCloud:CloudRequestFactory=nodeCloudRequest){
  const url=new URL(req.url??'/',origin);
  const match=/^\/api\/cloud-proxy\/([^/]+)(\/api\/cloud\/(?:[^/]+|admin\/(?:user|users|release|releases)|releases\/(?:latest|download)))$/.exec(url.pathname);
  let base:URL;
  try{
    if(!match)throw new Error();
    base=new URL(decodeURIComponent(match[1]));
    if(!['http:','https:'].includes(base.protocol)||base.username||base.password||base.search||base.hash)throw new Error();
    if(req.headers.origin&&req.headers.origin!==origin||req.headers['sec-fetch-site']==='cross-site')throw new Error();
    if(!['GET','POST','PUT','PATCH','DELETE','HEAD'].includes(req.method??''))throw new Error();
  }catch{res.writeHead(403);res.end();return;}

  const server=base.origin+base.pathname.replace(/\/+$/,'');
  const target=new URL(server+match![2]+url.search);
  const cookieName='mydoc-cloud-'+createHash('sha256').update(server).digest('hex');
  const cookiePath='/api/cloud-proxy/'+match![1]+'/';
  const cookie=req.headers.cookie?.split(';').map(value=>value.trim()).find(value=>value.startsWith(cookieName+'='));
  const headers:Record<string,string>={};
  // Forward only cloud credentials, never the desktop bootstrap cookie or IPC token.
  if(cookie)headers.Cookie='mydoc_session='+cookie.slice(cookieName.length+1);
  for(const name of ['content-type','x-mydoc-account','user-agent','x-mydoc-version','x-mydoc-platform','x-mydoc-arch','x-mydoc-filename','x-mydoc-notes']){
    const value=req.headers[name];if(typeof value==='string')headers[name]=value;
  }
  const fail=(error:NodeJS.ErrnoException)=>{
    clearTimeout(timer);
    if(res.destroyed)return;
    if(res.headersSent){res.destroy();return;}
    res.writeHead(502,{'Content-Type':'application/json'});
    res.end(JSON.stringify({error:cloudConnectionError(error,server),code:error.code??'UNKNOWN'}));
  };
  const timer=setTimeout(()=>{fail(Object.assign(new Error('upstream timeout'),{code:'ETIMEDOUT'}));upstream.cancel();},15000);
  const upstream=requestCloud(target,req.method??'GET',headers,response=>{
    // Only the handshake has a deadline; SSE remains open until the client closes.
    clearTimeout(timer);
    res.statusCode=response.statusCode??502;
    res.setHeader('Cache-Control','no-store');
    res.setHeader('X-Content-Type-Options','nosniff');
    res.setHeader('Content-Type',response.headers['content-type']??'application/octet-stream');
    if(response.headers['content-disposition'])res.setHeader('Content-Disposition',response.headers['content-disposition']);
    const cookies=(response.headers['set-cookie']??[]).filter(value=>value.startsWith('mydoc_session=')).map(value=>{
      const parts=value.split(';').map(part=>part.trim());
      const expiry=parts.filter(part=>/^(Max-Age|Expires)=/i.test(part));
      return cookieName+'='+parts[0].slice('mydoc_session='.length)+'; HttpOnly; SameSite=Strict; Path='+cookiePath+(expiry.length?'; '+expiry.join('; '):'');
    });
    if(cookies.length)res.setHeader('Set-Cookie',cookies);
    response.on('error',()=>res.destroy());
    response.pipe(res);
  });
  upstream.body.on('error',fail);
  const cancel=()=>{clearTimeout(timer);upstream.cancel();};
  res.on('close',cancel);
  req.on('aborted',cancel);
  req.pipe(upstream.body);
}
