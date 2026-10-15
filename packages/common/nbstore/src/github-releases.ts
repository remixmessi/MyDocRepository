import {validReleaseVersion} from './releases';

export const githubRepository = 'https://github.com/remixmessi/MyDocRepository';
export const githubReleasesPage = githubRepository + '/releases';
export const githubLatestReleaseApi = 'https://api.github.com/repos/remixmessi/MyDocRepository/releases/latest';
export interface GitHubInstaller {name:string;url:string;bytes?:number;sha256?:string;}
export interface GitHubRelease {version:string;notes:string;publishedAt?:number;url:string;installer?:GitHubInstaller;}

function repositoryUrl(value:unknown, path:string):value is string {
 if(typeof value!=='string')return false;
 try {const url=new URL(value);return url.protocol==='https:'&&url.host==='github.com'&&!url.username&&!url.password&&!url.search&&!url.hash&&url.pathname.startsWith('/remixmessi/MyDocRepository/releases/'+path);}
 catch{return false;}
}
/** Match explicit OS and architecture tokens; never guess an unlabelled package's architecture. */
function installerRank(name:string,platform:string,arch:string) {
 const tokens=name.toLowerCase().split(/[^a-z0-9]+/);
 const systems:Record<string,string[]>={darwin:['mac','macos','darwin','osx'],win32:['win','win32','windows'],linux:['linux']};
 const architectures:Record<string,string[]>={x64:['x64','amd64','x86_64'],arm64:['arm64','aarch64']};
 const hasArch=(target:string)=>(architectures[target]??[]).some(token=>token==='x86_64'?/(?:^|[^a-z0-9])x86_64(?:[^a-z0-9]|$)/i.test(name):tokens.includes(token));
 if(!systems[platform]?.some(token=>tokens.includes(token))||Object.entries(systems).some(([key,values])=>key!==platform&&values.some(token=>tokens.includes(token))))return 0;
 const universal=platform==='darwin'&&tokens.includes('universal');
 if(!['x64','arm64'].includes(arch)||!universal&&(!hasArch(arch)||hasArch(arch==='x64'?'arm64':'x64')))return 0;
 const extension=name.slice(name.lastIndexOf('.')+1).toLowerCase();
 const formats:Record<string,string[]>={darwin:['dmg','pkg','zip'],win32:['exe','msi'],linux:['appimage','deb','rpm']};
 const rank=formats[platform]?.indexOf(extension)??-1;
 return rank<0?0:10-rank+(universal?0:10);
}
export function parseGitHubRelease(value:unknown,platform:string,arch:string,allowMissingSizes=false):GitHubRelease {
 if(!value||typeof value!=='object')throw new Error('GitHub 返回的版本信息无效');
 const data=value as Record<string,unknown>;
 const version=typeof data.tag_name==='string'?data.tag_name.replace(/^v/i,''):'';
 if(!validReleaseVersion(version)||version.split('+')[0].includes('-')||data.draft!==false||data.prerelease!==false||!repositoryUrl(data.html_url,'tag/'))throw new Error('GitHub 未提供有效的正式版本信息');
 const installers=(Array.isArray(data.assets)?data.assets:[]).flatMap((item:unknown)=>{
  if(!item||typeof item!=='object')return [];
  const asset=item as Record<string,unknown>;
  if(typeof asset.name!=='string'||asset.state!=='uploaded'||!(allowMissingSizes&&asset.size===undefined)&&!(typeof asset.size==='number'&&Number.isSafeInteger(asset.size)&&asset.size>0)||!repositoryUrl(asset.browser_download_url,'download/'))return [];
  const rank=installerRank(asset.name,platform,arch);if(!rank)return [];
  return [{rank,name:asset.name,url:asset.browser_download_url,bytes:asset.size as number|undefined,sha256:typeof asset.digest==='string'&&/^sha256:[a-f0-9]{64}$/i.test(asset.digest)?asset.digest.slice(7):undefined}];
 }).sort((a,b)=>b.rank-a.rank||a.name.localeCompare(b.name));
 const date=typeof data.published_at==='string'?Date.parse(data.published_at):NaN;
 return {version,notes:typeof data.body==='string'?data.body:'',publishedAt:Number.isFinite(date)?date:undefined,url:data.html_url,installer:installers[0]};
}
export async function checkGitHubRelease(platform:string,arch:string,signal:AbortSignal,fallback?:()=>Promise<GitHubRelease|null>):Promise<GitHubRelease|null> {
 const response=await fetch(githubLatestReleaseApi,{headers:{Accept:'application/vnd.github+json'},signal,credentials:'omit',cache:'no-store'});
 if(response.status===404)return null;
 if(response.status===403||response.status===429){
  if(fallback)return fallback();
  throw new Error(response.status===429||response.headers.get('x-ratelimit-remaining')==='0'?'GitHub API 请求额度已用尽，请稍后重试或查看 Release 页面':'GitHub API 拒绝了请求（403），请检查网络或查看 Release 页面');
 }
 if(!response.ok)throw new Error('GitHub 版本检查失败（'+response.status+'），请稍后重试');
 return parseGitHubRelease(await response.json(),platform,arch);
}
