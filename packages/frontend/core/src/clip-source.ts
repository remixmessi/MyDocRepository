export interface WebClipSource { url:string; requestedUrl:string; title:string; clippedAt:number }
export function clipUrl(value:string) {
 let url:URL;try{url=new URL(value.trim());}catch{throw new Error('请输入完整的 HTTP/HTTPS 网页网址');}
 if(!['http:','https:'].includes(url.protocol)||url.username||url.password||url.port||value.length>8192)throw new Error('仅支持不含账号密码的标准 HTTP/HTTPS 网址');
 url.hash='';return url.href;
}
export function validWebClip(value:unknown):value is WebClipSource {
 if(!value||typeof value!=='object')return false;const v=value as WebClipSource;
 try{return typeof v.url==='string'&&clipUrl(v.url)===v.url&&typeof v.requestedUrl==='string'&&clipUrl(v.requestedUrl)===v.requestedUrl&&typeof v.title==='string'&&v.title.length<=500&&typeof v.clippedAt==='number'&&Number.isFinite(v.clippedAt)&&v.clippedAt>0;}catch{return false;}
}
