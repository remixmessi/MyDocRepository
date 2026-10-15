import {validReleaseVersion} from '../../../../common/nbstore/src/releases';
import {githubRepository,parseGitHubRelease} from '../../../../common/nbstore/src/github-releases';
interface UpdateFile {url:string;sha512:string;size?:number;}
/** Filter before electron-updater caches the event's info: its default file selection may fall back to the wrong architecture. */
export function githubUpdateFiles(info:{version:string;files:UpdateFile[]},platform:string,arch:string):UpdateFile[] {
 const invalid=()=>Object.assign(new Error('GitHub release has no matching, verifiable update package'),{code:'MYDOC_GITHUB_PACKAGE_UNAVAILABLE'});
 if(!validReleaseVersion(info.version)||!Array.isArray(info.files)||!['darwin','win32'].includes(platform))throw invalid();
 const extension=platform==='darwin'?'.zip':'.exe';
 const files=info.files.filter(file=>{
  if(!file||typeof file.url!=='string'||file.url!==file.url.split('/').pop()||file.url.includes('\\')||!file.url.toLowerCase().endsWith(extension)||!/^[-\w. ]+$/.test(file.url)||!/^[-A-Za-z0-9+/]{86}==$/.test(file.sha512??''))return false;
  const release=parseGitHubRelease({tag_name:info.version,draft:false,prerelease:false,html_url:githubRepository+'/releases/tag/'+info.version,assets:[{name:file.url,state:'uploaded',size:file.size,browser_download_url:githubRepository+'/releases/download/'+info.version+'/'+encodeURIComponent(file.url)}]},platform,arch,true);
  return !!release.installer;
 });
 if(!files.length)throw invalid();
 return files;
}
