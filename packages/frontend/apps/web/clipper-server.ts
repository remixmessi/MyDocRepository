import {lookup} from 'node:dns/promises';
import {isIP} from 'node:net';
import {request as httpRequest,type IncomingMessage,type ServerResponse,type RequestOptions} from 'node:http';
import {request as httpsRequest} from 'node:https';
import {createGunzip,createBrotliDecompress,createInflate} from 'node:zlib';
import type {Plugin} from 'vite';
export const CLIP_LIMITS={html:3*1024*1024,image:20*1024*1024,timeout:15000,redirects:5};
export function publicClipAddress(address:string) {
 if(isIP(address)===4){const [a,b,c]=address.split('.').map(Number);return !(a===0||a===10||a===127||a>=224||a===169&&b===254||a===172&&b>=16&&b<=31||a===192&&(b===168||b===0||b===2)||a===100&&b>=64&&b<=127||a===198&&(b===18||b===19||b===51&&c===100)||a===203&&b===0&&c===113);}
 return isIP(address)===6&&/^[23]/i.test(address)&&!/^2001:(?:db8|0):/i.test(address)&&!/^2002:/i.test(address);
}
export function clipRequestUrl(value:unknown) {if(typeof value!=='string'||value.length>8192)throw new Error('请输入有效网页网址');const url=new URL(value);if(!['http:','https:'].includes(url.protocol)||url.username||url.password||url.port)throw new Error('仅支持不含账号密码的标准 HTTP/HTTPS 网址');url.hash='';return url;}
export async function clipAddress(url:URL,resolve=lookup) {const host=url.hostname.replace(/^\[|\]$/g,'');const addresses=isIP(host)?[{address:host,family:isIP(host)}]:await resolve(host,{all:true});if(!addresses.length||addresses.some(ip=>!publicClipAddress(ip.address)))throw new Error('无法读取本机或内网地址，请选择公开网页');return addresses[0];}
export async function readPublicClip(value:string,kind:'page'|'image',signal:AbortSignal) {
 let url=clipRequestUrl(value);const limit=kind==='page'?CLIP_LIMITS.html:CLIP_LIMITS.image;
 for(let redirects=0;redirects<=CLIP_LIMITS.redirects;redirects++){
  if(signal.aborted)throw new Error('网页读取已停止');const ip=await clipAddress(url);if(signal.aborted)throw new Error('网页读取已停止');
  const result=await new Promise<{redirect?:string;bytes?:Buffer;mime?:string}>((resolve,reject)=>{
   const request=(url.protocol==='https:'?httpsRequest:httpRequest)(url,{signal,autoSelectFamily:false,lookup:(_host:string,_options:unknown,done:(error:Error|null,address:string,family:number)=>void)=>done(null,ip.address,ip.family),headers:{'User-Agent':'MyDoc-Clipper/1.0','Accept':kind==='page'?'text/html,application/xhtml+xml':'image/png,image/jpeg,image/gif,image/webp','Accept-Encoding':'identity'}} as RequestOptions & {autoSelectFamily:boolean},response=>{
    const status=response.statusCode??0;if([301,302,303,307,308].includes(status)){const location=response.headers.location;response.destroy();if(!location)reject(new Error('网页跳转地址缺失'));else resolve({redirect:location});return;}
    if(status<200||status>=300){response.destroy();reject(new Error(`网页返回 ${status}，可能需要登录或不允许读取`));return;}
    const mime=response.headers['content-type']??'';if(kind==='page'&&!/^(text\/html|application\/xhtml\+xml)(;|$)/i.test(mime)){response.destroy();reject(new Error('该网址没有返回 HTML 网页'));return;}
    if(Number(response.headers['content-length']??0)>limit){response.destroy();reject(new Error('网页或图片超过大小限制'));return;}
    const encoding=response.headers['content-encoding'];const decoder=encoding==='gzip'?createGunzip():encoding==='br'?createBrotliDecompress():encoding==='deflate'?createInflate():null;
    if(encoding&&encoding!=='identity'&&!decoder){response.destroy();reject(new Error('网页压缩格式不支持'));return;}
    const stream=decoder?response.pipe(decoder):response;const chunks:Buffer[]=[];let size=0;
    stream.on('data',(chunk:Buffer)=>{size+=chunk.length;if(size>limit){stream.destroy();response.destroy();reject(new Error('网页或图片超过大小限制'));}else chunks.push(chunk);});stream.on('error',reject);response.on('error',reject);stream.on('end',()=>resolve({bytes:Buffer.concat(chunks),mime}));
   });request.on('error',reject);
   request.end();
  });
  if(result.redirect){url=clipRequestUrl(new URL(result.redirect,url).href);continue;}
  return {url:url.href,bytes:result.bytes!,mime:result.mime!};
 }
 throw new Error('网页跳转次数过多');
}
export async function clipperHandler(req:IncomingMessage,res:ServerResponse) {
 if(req.method!=='POST'){res.writeHead(405,{'Allow':'POST'});res.end();return;}
 try{
  const origin=req.headers.origin;if(!origin||new URL(origin).host!==req.headers.host||!['localhost','127.0.0.1','[::1]'].includes(new URL(origin).hostname)||!req.headers['content-type']?.startsWith('application/json'))throw new Error('网页读取请求来源无效');
  let body='';for await(const chunk of req){body+=chunk.toString();if(body.length>10000)throw new Error('读取请求过大');}const input=JSON.parse(body);if(!['page','image'].includes(input.kind))throw new Error('读取类型无效');
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),CLIP_LIMITS.timeout);const abort=()=>{if(!res.writableEnded)controller.abort();};res.on('close',abort);
  try{const result=await readPublicClip(input.url,input.kind,controller.signal);if(input.kind==='image'){if(!/^image\/(png|jpeg|gif|webp)(;|$)/i.test(result.mime))throw new Error('图片格式不支持，请使用远程链接');res.writeHead(200,{'Content-Type':result.mime,'Cache-Control':'no-store'});res.end(result.bytes);}else{const prefix=result.bytes.subarray(0,4096).toString(),charset=/charset\s*=\s*["']?([\w-]+)/i.exec(result.mime)?.[1]??/<meta[^>]*charset\s*=\s*["']?([\w-]+)/i.exec(prefix)?.[1]??'utf-8';let html:string;try{html=new TextDecoder(charset).decode(result.bytes);}catch{html=result.bytes.toString('utf8');}res.writeHead(200,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify({url:result.url,html}));}}finally{clearTimeout(timer);res.off('close',abort);}
 }catch(error){if(!res.headersSent&&!res.destroyed){res.writeHead(400,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify({error:error instanceof Error&&error.name==='AbortError'?'网页读取超时，请重试':error instanceof Error?error.message:'网页读取失败，请重试'}));}}
}
export function clipperServer():Plugin {return {name:'mydoc-clipper',configureServer(server){server.middlewares.use('/api/clipper',(req,res)=>void clipperHandler(req,res));},configurePreviewServer(server){server.middlewares.use('/api/clipper',(req,res)=>void clipperHandler(req,res));}};}
