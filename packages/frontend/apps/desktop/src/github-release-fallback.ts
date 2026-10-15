import {githubReleasesPage,parseGitHubRelease,type GitHubRelease} from '../../../../common/nbstore/src/github-releases';

/** Read only fixed public GitHub pages; never accept a renderer-supplied upstream URL. */
export function createGitHubReleaseFallback(request:typeof fetch=fetch) {
 const cache=new Map<string,{until:number;value:GitHubRelease|null}>();
 const pending=new Map<string,Promise<GitHubRelease|null>>();
 return async(platform:string,arch:string)=>{
  if(!['darwin','win32','linux'].includes(platform)||!['arm64','x64'].includes(arch))throw new Error('客户端系统或架构无效');
  const key=platform+'/'+arch,cached=cache.get(key);if(cached&&cached.until>Date.now())return cached.value;
  const active=pending.get(key);if(active)return active;
  const task=(async()=>{
   const signal=AbortSignal.timeout(15000);
   const get=async(url:string)=>request(url,{signal,credentials:'omit',headers:{Accept:'text/html'},redirect:'manual'});
   // Resolve /latest ourselves so a redirect can never send the request outside this repository.
   const latest=await get(githubReleasesPage+'/latest');
   if(latest.status===404)return null;
   const location=latest.headers.get('location');
   if(![301,302,303,307,308].includes(latest.status)||!location)throw new Error('GitHub 页面检查失败（'+latest.status+'），请打开 Release 页面查看');
   const target=new URL(location,githubReleasesPage+'/latest');
   const prefix='/remixmessi/MyDocRepository/releases/tag/';
   if(target.origin!=='https://github.com'||!target.pathname.startsWith(prefix)||target.search||target.hash)throw new Error('GitHub 未返回有效的正式版本页面');
   const tag=decodeURIComponent(target.pathname.slice(prefix.length));
   // Validate the version before making another request.
   const base={tag_name:tag,draft:false,prerelease:false,html_url:target.href};
   parseGitHubRelease({...base,assets:[]},platform,arch);
   const assets=await get(githubReleasesPage+'/expanded_assets/'+encodeURIComponent(tag));
   if(!assets.ok)throw new Error('GitHub 安装包列表读取失败（'+assets.status+'），请打开 Release 页面查看');
   const html=await assets.text();
   const downloadPrefix='/remixmessi/MyDocRepository/releases/download/'+encodeURIComponent(tag)+'/';
   const urls=new Set<string>();
   for(const match of html.matchAll(/\bhref\s*=\s*["']([^"']+)["']/gi)){
    const path=match[1].replace(/&amp;/g,'&');
    if(!path.startsWith(downloadPrefix))continue;
    const name=path.slice(downloadPrefix.length);if(!name||name.includes('/'))continue;
    urls.add('https://github.com'+path);
   }
   const result=parseGitHubRelease({...base,assets:[...urls].map(url=>({name:decodeURIComponent(url.slice(url.lastIndexOf('/')+1)),state:'uploaded',browser_download_url:url}))},platform,arch,true);
   // The HTML fallback does not provide authoritative sizes, notes or checksums.
   return result;
  })();
  pending.set(key,task);
  try {const value=await task;cache.set(key,{value,until:Date.now()+5*60_000});return value;}
  finally{pending.delete(key);}
 };
}
