import {readFileSync,existsSync,writeFileSync,renameSync} from 'node:fs';
import {join} from 'node:path';
import {app} from 'electron';
import {autoUpdater} from 'electron-updater';
import {DesktopUpdates} from './updates';
import {githubUpdateFiles} from './github-update-info';
import {normalizeUpdateServer,serverUpdateFeed,serverUpdateFiles} from '../../../../common/nbstore/src/server-updates';
import type {DesktopUpdateState} from '../../../../common/nbstore/src/updates';

export const githubUpdateFeed={provider:'github' as const,owner:'remixmessi',repo:'MyDocRepository',releaseType:'release' as const};
/** Only expose actionable messages, never raw provider errors or credentials. */
export function githubUpdateError(error:unknown):string|undefined {
 const code=error&&typeof error==='object'&&'code' in error?String(error.code):'';
 if(code==='MYDOC_GITHUB_PACKAGE_UNAVAILABLE')return '此版本缺少适用于当前系统与架构的更新包或校验信息，请联系发布者。macOS 需要 ZIP，Windows 需要 EXE。';
 if(code==='ERR_UPDATER_CHANNEL_FILE_NOT_FOUND')return '此 GitHub Release 缺少更新清单，请发布 latest-mac.yml（macOS）或 latest.yml（Windows）及对应安装包。';
 if(['ERR_UPDATER_NO_PUBLISHED_VERSIONS','ERR_UPDATER_LATEST_VERSION_NOT_FOUND','ERR_UPDATER_INVALID_RELEASE_FEED'].includes(code))return '无法读取 GitHub 正式版本，请确认已发布公开 Release 并检查网络后重试。';
 if(['ERR_UPDATER_NO_ZIP','ERR_UPDATER_ZIP_FILE_NOT_FOUND'].includes(code))return '此版本缺少 macOS 更新 ZIP，请在同一 Release 上传打包生成的 ZIP 和 latest-mac.yml。';
 if(['ERR_UPDATER_INVALID_SIGNATURE','ERR_UPDATER_INVALID_CHECKSUM','ERR_CHECKSUM_MISMATCH'].includes(code))return '更新包校验失败，已阻止安装，请重新下载或联系发布者。';
 return undefined;
}
export function serverUpdateError(error:unknown):string|undefined {
 const details=error&&typeof error==='object'?error as {code?:unknown;statusCode?:unknown}:{};
 const code=String(details.code??''),status=Number(details.statusCode??(/^HTTP_ERROR_(\d+)$/.exec(code)?.[1]));
 if(code==='MYDOC_SERVER_PACKAGE_UNAVAILABLE')return '服务器返回的更新包与当前系统或架构不匹配，或缺少有效校验信息。';
 if(status===401||status===403)return '服务器更新接口拒绝访问（HTTP '+status+'）。请管理员升级服务器程序并允许匿名访问 /api/cloud/releases/update/；应用内检查更新无需登录。';
 if(code==='ERR_UPDATER_CHANNEL_FILE_NOT_FOUND'||status===404)return '服务器暂无匹配的正式更新包，或服务版本过旧。请管理员上传对应架构的 macOS ZIP / Windows EXE 并升级服务器。';
 if(status===409)return '服务器最新版本的更新包尚未就绪或校验失败，请管理员上传正确的 macOS ZIP / Windows EXE。';
 if(status>=500&&status<=599)return '服务器更新服务异常（HTTP '+status+'），请管理员检查服务日志后重试。';
 if(code==='ERR_UPDATER_INVALID_UPDATE_INFO')return '服务器返回的更新清单无效，请管理员检查更新接口及反向代理配置。';
 if(['ECONNREFUSED','ENOTFOUND','EAI_AGAIN','ETIMEDOUT','ERR_CONNECTION_REFUSED','ERR_NAME_NOT_RESOLVED','ERR_CONNECTION_TIMED_OUT'].includes(code))return '无法连接更新服务器，请检查服务器地址、端口及网络连接后重试。';
 if(['ERR_UPDATER_INVALID_SIGNATURE','ERR_UPDATER_INVALID_CHECKSUM','ERR_CHECKSUM_MISMATCH'].includes(code))return '更新包校验失败，已阻止安装，请重新下载或联系管理员。';
 return undefined;
}
export function createDesktopUpdates(options:{prepare:()=>Promise<void>;release:()=>void;notify:(state:DesktopUpdateState)=>void}){
 let config:{enabled?:boolean;channel?:string}={};try{config=JSON.parse(readFileSync(join(__dirname,'release-config.json'),'utf8'));}catch{}
 const supported=app.isPackaged&&['darwin','win32'].includes(process.platform);
 const configured=supported&&config.enabled===true&&existsSync(join(process.resourcesPath,'app-update.yml'));
 const githubDisabledReason=app.isPackaged?'当前系统不支持应用内更新。':'开发运行无法安装更新，请使用已打包的 macOS 或 Windows 客户端。';
 const settingsFile=join(app.getPath('userData'),'update-settings.json');let automaticChecks=true,savedSource:DesktopUpdateState['source'],savedServer:string|undefined;try{const settings=JSON.parse(readFileSync(settingsFile,'utf8'));automaticChecks=settings.automaticChecks!==false;if(settings.source==='server'){savedServer=normalizeUpdateServer(settings.serverUrl);savedSource='server';}else if(settings.source==='github')savedSource='github';}catch{}
 const source=savedSource??(configured?'configured':'github');
 const writeSettings=(value:boolean,nextSource:DesktopUpdateState['source'],serverUrl?:string)=>{writeFileSync(settingsFile+'.tmp',JSON.stringify({automaticChecks:value,source:nextSource,serverUrl}),{mode:0o600});renameSync(settingsFile+'.tmp',settingsFile);};
 autoUpdater.autoDownload=false;autoUpdater.autoInstallOnAppQuit=false;autoUpdater.autoRunAppAfterInstall=true;autoUpdater.allowDowngrade=false;autoUpdater.allowPrerelease=false;
 // User-entered server addresses are normalized in main; package URLs, paths and commands are never accepted from IPC.
 autoUpdater.logger=null;autoUpdater.disableDifferentialDownload=source==='server';
 if(supported&&source==='server'){autoUpdater.setFeedURL({provider:'generic',url:serverUpdateFeed(savedServer,process.platform,process.arch),useMultipleRangeRequest:false});autoUpdater.disableDifferentialDownload=true;}else if(supported&&source==='github')autoUpdater.setFeedURL(githubUpdateFeed);
 const updates:DesktopUpdates=new DesktopUpdates({...options,version:app.getVersion(),enabled:supported,source,serverUrl:source==='server'?savedServer:undefined,channel:source==='server'?'服务器更新':source==='configured'?config.channel??'稳定版':'GitHub 正式版',disabledReason:supported?undefined:githubDisabledReason,automaticChecks,
  driver:{check:()=>autoUpdater.checkForUpdates(),download:()=>autoUpdater.downloadUpdate(),install:()=>autoUpdater.quitAndInstall(false,true)},
  errorMessage:error=>updates.status.source==='github'?githubUpdateError(error):updates.status.source==='server'?serverUpdateError(error):undefined,
  saveSettings(value){writeSettings(value,updates.status.source,updates.status.serverUrl);},
 });
 const available=(info:import('electron-updater').UpdateInfo)=>{if(updates.status.source==='github')info.files.splice(0,info.files.length,...githubUpdateFiles(info,process.platform,process.arch));if(updates.status.source==='server')serverUpdateFiles(info,process.platform,process.arch);updates.available(info.version,info.releaseDate,typeof info.releaseNotes==='string'?info.releaseNotes:info.releaseNotes?.map(note=>note.version+'\n'+note.note).join('\n\n'));};
 const unavailable=()=>updates.unavailable();
 const progress=(p:{percent:number;transferred:number;total:number})=>updates.progress(p.percent,p.transferred,p.total);
 const downloaded=()=>updates.downloaded();
 const failed=(error:Error)=>updates.failed(undefined,error);
 autoUpdater.on('update-available',available);autoUpdater.on('update-not-available',unavailable);autoUpdater.on('download-progress',progress);autoUpdater.on('update-downloaded',downloaded);autoUpdater.on('error',failed);
 let startup:ReturnType<typeof setTimeout>|undefined,timer:ReturnType<typeof setInterval>|undefined;
 if(supported){startup=setTimeout(()=>{if(updates.status.automaticChecks)void updates.check();},30000);timer=setInterval(()=>{if(updates.status.automaticChecks)void updates.check();},6*3600000);startup.unref();timer.unref();}
 return {updates,async checkGitHub(){
  if(!updates.canChangeChannel)return updates.status;
  if(updates.status.source!=='github'){
   writeSettings(updates.status.automaticChecks,'github');if(supported){autoUpdater.setFeedURL(githubUpdateFeed);autoUpdater.disableDifferentialDownload=false;}
   updates.selectChannel({source:'github',channel:'GitHub 正式版',enabled:supported,disabledReason:supported?undefined:githubDisabledReason});
  }
  return updates.check();
 },async checkServer(address:unknown){
  if(!updates.canChangeChannel)return updates.status;
  const serverUrl=normalizeUpdateServer(address);
  if(updates.status.source!=='server'||updates.status.serverUrl!==serverUrl){
   const feed=serverUpdateFeed(serverUrl,process.platform,process.arch);writeSettings(updates.status.automaticChecks,'server',serverUrl);
   if(supported){autoUpdater.setFeedURL({provider:'generic',url:feed,useMultipleRangeRequest:false});autoUpdater.disableDifferentialDownload=true;}
   updates.selectChannel({source:'server',serverUrl,channel:'服务器更新',enabled:supported,disabledReason:supported?undefined:githubDisabledReason});
  }
  return updates.check();
 },stop(){if(startup)clearTimeout(startup);if(timer)clearInterval(timer);autoUpdater.removeListener('update-available',available);autoUpdater.removeListener('update-not-available',unavailable);autoUpdater.removeListener('download-progress',progress);autoUpdater.removeListener('update-downloaded',downloaded);autoUpdater.removeListener('error',failed);}};
}
