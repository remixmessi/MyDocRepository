import {readFileSync,existsSync,writeFileSync,renameSync} from 'node:fs';import {join} from 'node:path';
import {app} from 'electron';import {autoUpdater} from 'electron-updater';
import {DesktopUpdates} from './updates';import type {DesktopUpdateState} from '../../../../common/nbstore/src/updates';
export function createDesktopUpdates(options:{prepare:()=>Promise<void>;release:()=>void;notify:(state:DesktopUpdateState)=>void}){
 let config:{enabled?:boolean;channel?:string}={};try{config=JSON.parse(readFileSync(join(__dirname,'release-config.json'),'utf8'));}catch{}
 const enabled=app.isPackaged&&config.enabled===true&&existsSync(join(process.resourcesPath,'app-update.yml'))&&['darwin','win32'].includes(process.platform);
 const settingsFile=join(app.getPath('userData'),'update-settings.json');let automaticChecks=true;try{automaticChecks=JSON.parse(readFileSync(settingsFile,'utf8')).automaticChecks!==false;}catch{}
 autoUpdater.autoDownload=false;autoUpdater.autoInstallOnAppQuit=false;autoUpdater.autoRunAppAfterInstall=true;autoUpdater.allowDowngrade=false;autoUpdater.allowPrerelease=false;
 // No renderer-supplied feed URL. The signed build owns app-update.yml.
 autoUpdater.logger=null;
 const updates=new DesktopUpdates({...options,version:app.getVersion(),enabled,channel:enabled?config.channel??'稳定版':'未配置',disabledReason:enabled?undefined:'更新渠道未配置。当前版本不会检查、下载或安装在线更新。',automaticChecks,
  driver:enabled?{check:()=>autoUpdater.checkForUpdates(),download:()=>autoUpdater.downloadUpdate(),install:()=>autoUpdater.quitAndInstall(false,true)}:undefined,
  saveSettings(value){writeFileSync(settingsFile+'.tmp',JSON.stringify({automaticChecks:value}),{mode:0o600});renameSync(settingsFile+'.tmp',settingsFile);},
 });
 autoUpdater.on('update-available',info=>updates.available(info.version,info.releaseDate));autoUpdater.on('update-not-available',()=>updates.unavailable());autoUpdater.on('download-progress',p=>updates.progress(p.percent,p.transferred,p.total));autoUpdater.on('update-downloaded',()=>updates.downloaded());autoUpdater.on('error',()=>updates.failed());
 let startup:ReturnType<typeof setTimeout>|undefined,timer:ReturnType<typeof setInterval>|undefined;
 if(enabled){startup=setTimeout(()=>{if(updates.status.automaticChecks)void updates.check();},30000);timer=setInterval(()=>{if(updates.status.automaticChecks)void updates.check();},6*3600000);startup.unref();timer.unref();}
 return {updates,stop(){if(startup)clearTimeout(startup);if(timer)clearInterval(timer);}};
}
