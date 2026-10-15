import {validReleaseVersion} from './releases';
/** Normalize only the address explicitly entered in the update settings. */
export function normalizeUpdateServer(value:unknown){
 if(typeof value!=='string'||!value.trim()||value.length>2048)throw new Error('请填写更新服务器地址');
 let url:URL;try{url=new URL(/^[a-z][a-z\d+.-]*:\/\//i.test(value.trim())?value.trim():'http://'+value.trim());}catch{throw new Error('更新服务器地址无效');}
 if(!['http:','https:'].includes(url.protocol)||url.username||url.password||url.search||url.hash)throw new Error('请填写 HTTP / HTTPS 服务器地址，不含用户名、查询参数或锚点');
 return url.origin+url.pathname.replace(/\/+$/,'');
}
export function serverUpdateFeed(server:unknown,platform:string,arch:string){
 if(!['darwin','win32'].includes(platform)||!['arm64','x64'].includes(arch))throw new Error('当前系统或架构不支持应用内更新');
 return normalizeUpdateServer(server)+'/api/cloud/releases/update/'+platform+'/'+arch+'/';
}
interface FileInfo {url:string;sha512:string;size?:number;}
export function serverUpdateFiles(info:{version:string;files:FileInfo[];mydocPlatform?:unknown;mydocArch?:unknown},platform:string,arch:string){
 if(!validReleaseVersion(info.version)||info.version.split('+')[0].includes('-')||info.mydocPlatform!==platform||info.mydocArch!==arch||!Array.isArray(info.files)||info.files.length!==1)throw Object.assign(new Error('更新包系统或架构不匹配'),{code:'MYDOC_SERVER_PACKAGE_UNAVAILABLE'});
 const file=info.files[0],extension=platform==='darwin'?'zip':'exe';
 if(!file||!new RegExp('^[a-f0-9-]{36}\\.'+extension+'$').test(file.url)||!/^[A-Za-z0-9+/]{86}==$/.test(file.sha512)||!Number.isSafeInteger(file.size)||(file.size??0)<=0)throw Object.assign(new Error('更新包格式或校验信息无效'),{code:'MYDOC_SERVER_PACKAGE_UNAVAILABLE'});
 return info.files;
}
