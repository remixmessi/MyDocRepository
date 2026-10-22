import {upgradeCloudDatabase,LOCAL_OFFICE_BACKUP} from '../../web/server/schema';
import {spawn} from 'node:child_process';
import {createDesktopUpdates} from './update-runtime';
import {DocumentLinkQueue,commandLineLinks} from './deep-links';
import {OpenFileQueue,commandLineFiles} from './open-files';
import {app,BrowserWindow,Menu,ipcMain,dialog,shell,net} from 'electron';
import {dirname,join,resolve} from 'node:path';
import {randomUUID} from 'node:crypto';
import {mkdirSync} from 'node:fs';
import {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {connectDesktopDatabase} from './sqlite-driver';
import {startDesktopServer} from './server';
import {electronCloudRequest} from './cloud-transport';
import {desktopDatabasePath,migrateLegacyDesktopDatabase,safeExternalUrl,validStorageRequest} from './policy';
import type {DesktopStorageReply} from '../../../../common/nbstore/src/desktop';
app.setName('MyDoc');
const profile=!app.isPackaged&&process.env.MYDOC_DESKTOP_PROFILE?resolve(process.env.MYDOC_DESKTOP_PROFILE):join(app.getPath('appData'),app.isPackaged?'MyDoc':'MyDoc Development');
mkdirSync(profile,{recursive:true});app.setPath('userData',profile);
let database:SQLiteProjectDatabase|undefined,server:Awaited<ReturnType<typeof startDesktopServer>>|undefined,window:BrowserWindow|undefined;
let allowingQuit=false,closing=false,updateOwnsClose=false,updateOwnsDatabaseLock=false;
let updater:ReturnType<typeof createDesktopUpdates>|undefined;
const openFiles=new OpenFileQueue();
const documentLinks=new DocumentLinkQueue();
function enqueueLink(url:string){if(!documentLinks.enqueue(url))void app.whenReady().then(()=>dialog.showMessageBox({type:'info',message:'文档链接等待队列已满，请完成当前操作后重试。'}));focusWindow();window?.webContents.send('mydoc:document-link-available');}
app.on('open-url',(event,url)=>{event.preventDefault();enqueueLink(url);});
function focusWindow(){if(window&&!window.isDestroyed()){if(window.isMinimized())window.restore();window.show();window.focus();}else if(server)makeWindow();}
function enqueueFiles(paths:string[]){if(!openFiles.enqueue(paths)&&paths.length&&openFiles.pending>=8){void app.whenReady().then(()=>dialog.showMessageBox({type:'info',message:'文件已在等待导入，或等待队列已满。请完成当前导入后重试。'}));}focusWindow();window?.webContents.send('mydoc:open-files-available');}
app.on('open-file',(event,path)=>{event.preventDefault();enqueueFiles([path]);});
const pendingClose=new Map<string,{resolve:(result:{ok:boolean;error?:string})=>void;sender:number}>();
const appIcon=join(__dirname,'icon.png');
function trusted(event:Electron.IpcMainInvokeEvent|Electron.IpcMainEvent){return !!window&&event.sender===window.webContents&&event.senderFrame===event.sender.mainFrame&&!!server&&new URL(event.senderFrame.url).origin===server.origin;}
ipcMain.handle('mydoc:update-status',event=>{if(!trusted(event)||!updater)throw new Error('更新请求无效');return updater.updates.status;});
ipcMain.handle('mydoc:update-action',(event,action:unknown)=>{if(!trusted(event)||!updater||typeof action!=='string'||!['check','check-github','download','install'].includes(action))throw new Error('更新请求无效');return action==='check-github'?updater.checkGitHub():action==='check'?updater.updates.check():action==='download'?updater.updates.download():updater.updates.install();});
ipcMain.handle('mydoc:update-server',(event,address:unknown)=>{if(!trusted(event)||!updater)throw new Error('更新请求无效');return updater.checkServer(address);});
ipcMain.handle('mydoc:update-settings',(event,value:unknown)=>{if(!trusted(event)||!updater)throw new Error('更新请求无效');return updater.updates.configure(value);});
ipcMain.handle('mydoc:storage',async(event,input):Promise<DesktopStorageReply>=>{
 if(!trusted(event)||!validStorageRequest(input)||!server)throw new Error('桌面存储请求无效');
 const response=await fetch(server.origin+'/api/storage'+input.path,{method:input.method,headers:{Authorization:'Bearer '+server.token,...(input.generation?{'X-Mydoc-Generation':input.generation}:{}),...(input.officeGeneration?{'X-Mydoc-Office-Generation':input.officeGeneration}:{}),...(input.contentType?{'Content-Type':input.contentType}:{})},body:input.body?.slice().buffer});
 return {status:response.status,officeGeneration:response.headers.get('x-mydoc-office-generation')??undefined,contentType:response.headers.get('content-type')??'application/octet-stream',body:new Uint8Array(await response.arrayBuffer())};
});
ipcMain.handle('mydoc:choose-backup-directory',async event=>{if(!trusted(event)||!server||!window)throw new Error('备份目录请求无效');const result=await dialog.showOpenDialog(window,{title:'选择备份目录',properties:['openDirectory','createDirectory']});return result.canceled?undefined:server.backups.setDirectory(result.filePaths[0]);});
ipcMain.handle('mydoc:open-backup-directory',async event=>{if(!trusted(event)||!server)throw new Error('备份目录请求无效');mkdirSync(server.backups.settings.directory,{recursive:true});const error=await shell.openPath(server.backups.settings.directory);if(error)throw new Error('无法打开备份目录');});
ipcMain.handle('mydoc:take-document-link',event=>{if(!trusted(event))throw new Error('文档打开请求无效');return documentLinks.take();});
ipcMain.handle('mydoc:ack-document-link',(event,id:unknown)=>{if(!trusted(event))throw new Error('文档打开请求无效');documentLinks.acknowledge(id);if(documentLinks.pending)window?.webContents.send('mydoc:document-link-available');});
ipcMain.handle('mydoc:take-open-files',event=>{if(!trusted(event))throw new Error('文件打开请求无效');return openFiles.take();});
ipcMain.handle('mydoc:ack-open-files',(event,id:unknown)=>{if(!trusted(event))throw new Error('文件打开请求无效');openFiles.acknowledge(id);if(openFiles.pending)window?.webContents.send('mydoc:open-files-available');});
ipcMain.on('mydoc:close-result',(event,result:unknown)=>{if(!trusted(event)||!result||typeof result!=='object')return;const r=result as {id?:unknown;ok?:unknown;error?:unknown};if(typeof r.id!=='string'||typeof r.ok!=='boolean')return;const pending=pendingClose.get(r.id);if(pending&&pending.sender===event.sender.id)pending.resolve({ok:r.ok,error:typeof r.error==='string'?r.error.slice(0,2000):undefined});});
async function saveBeforeClose(lock=false){
 if(server?.backups.running){await dialog.showMessageBox({type:'info',message:'备份或恢复正在进行，请完成后再关闭。'});return false;}
 const target=window;if(!target||target.isDestroyed())return true;
 if(target.webContents.isLoading()){await dialog.showMessageBox(target,{type:'info',message:'应用正在打开，请稍后再关闭。'});return false;}
 const id=randomUUID();let timer:ReturnType<typeof setTimeout>;
 const result=await new Promise<{ok:boolean;error?:string}>(resolve=>{pendingClose.set(id,{resolve,sender:target.webContents.id});timer=setTimeout(()=>resolve({ok:false,error:'等待保存超时，请返回文档检查保存状态后重试。'}),15000);target.webContents.send(lock?'mydoc:prepare-update':'mydoc:prepare-close',id);});
 clearTimeout(timer!);pendingClose.delete(id);
 if(!result.ok){await dialog.showMessageBox(target,{type:'error',title:'尚未完成保存',message:'已保留窗口，请完成保存后再关闭。',detail:result.error??'保存失败，请重试。',buttons:['返回文档']});return false;}
 return true;
}
async function requestQuit(){if(allowingQuit||closing)return;closing=true;try{if(await saveBeforeClose()){allowingQuit=true;app.quit();}}finally{closing=false;}}
async function uninstallMac(){
 if(!app.isPackaged||process.platform!=='darwin'||closing)return;
 const application=resolve(app.getPath('exe'),'../../..');
 if(!application.endsWith('/MyDoc.app')){dialog.showErrorBox('无法自动卸载','请先将应用恢复为 MyDoc.app 名称，再使用卸载入口。');return;}
 const result=await dialog.showMessageBox({type:'warning',title:'卸载 MyDoc',message:'卸载 MyDoc 并删除本机全部数据？',detail:'将删除应用、本机数据库、附件、内部备份和设置。此操作不可撤销。自行导出的文件和服务端数据保留。',buttons:['取消','卸载并删除本机数据'],defaultId:0,cancelId:0,noLink:true});
 if(result.response!==1||closing)return;
 closing=true;
 try{
  if(!await saveBeforeClose(true))return;
  const child=spawn('/bin/sh',[join(process.resourcesPath,'uninstall-macos.sh'),String(process.pid),app.getPath('userData'),application],{detached:true,stdio:'ignore'});
  await new Promise<void>((resolve,reject)=>{child.once('spawn',resolve);child.once('error',reject);});child.unref();
  allowingQuit=true;app.quit();
 }catch(error){window?.webContents.send('mydoc:update-unlock');dialog.showErrorBox('卸载未完成',error instanceof Error?error.message:String(error));}
 finally{if(!allowingQuit)window?.webContents.send('mydoc:update-unlock');closing=false;}
}
async function reload(){if(!closing&&await saveBeforeClose())window?.webContents.reload();}
function makeWindow(){
 if(!server)return;
 window=new BrowserWindow({width:1320,height:900,minWidth:760,minHeight:540,show:false,title:'MyDoc',icon:appIcon,backgroundColor:'#f5f5f5',webPreferences:{preload:join(__dirname,'preload.cjs'),contextIsolation:true,nodeIntegration:false,sandbox:true,webSecurity:true}});
 const current=window;
 current.webContents.setWindowOpenHandler(({url})=>{
  // Local PDF previews may open in a separate, unprivileged viewer.
  if(url.startsWith('blob:'+server!.origin+'/'))return {action:'allow',overrideBrowserWindowOptions:{title:'MyDoc 文件预览',webPreferences:{preload:'',contextIsolation:true,nodeIntegration:false,sandbox:true,webSecurity:true}}};
  if(/^mydoc:/i.test(url)){enqueueLink(url);return {action:'deny'};}const external=safeExternalUrl(url);if(external)void shell.openExternal(external);return {action:'deny'};
 });
 current.webContents.on('did-create-window',viewer=>{viewer.webContents.on('will-navigate',(event,url)=>{if(!url.startsWith('blob:'+server!.origin+'/'))event.preventDefault();});viewer.webContents.setWindowOpenHandler(()=>({action:'deny'}));});
 current.webContents.on('will-navigate',(event,url)=>{if(new URL(url).origin!==server?.origin){event.preventDefault();if(/^mydoc:/i.test(url)){enqueueLink(url);return;}const external=safeExternalUrl(url);if(external)void shell.openExternal(external);}});
 current.webContents.on('will-attach-webview',event=>event.preventDefault());
 current.webContents.on('before-input-event',(event,input)=>{if(input.type==='keyDown'&&((input.control||input.meta)&&input.key.toLowerCase()==='r'||input.key==='F5')){event.preventDefault();void reload();}});
 current.webContents.session.setPermissionRequestHandler((contents,permission,callback,details)=>callback(contents===current.webContents&&details.isMainFrame&&['clipboard-sanitized-write','clipboard-read','fullscreen'].includes(permission)));
 current.webContents.session.on('will-download',(_event,item)=>item.setSaveDialogOptions({title:'保存文件',defaultPath:join(app.getPath('downloads'),item.getFilename())}));
 current.once('ready-to-show',()=>current.show());
 current.webContents.on('did-finish-load',()=>{if(updater)current.webContents.send('mydoc:update-state',updater.updates.status);if(openFiles.pending)current.webContents.send('mydoc:open-files-available');if(documentLinks.pending)current.webContents.send('mydoc:document-link-available');});
 current.on('close',event=>{if(allowingQuit)return;event.preventDefault();if(closing)return;closing=true;void saveBeforeClose().then(ok=>{if(ok){if(process.platform==='darwin')current.destroy();else{allowingQuit=true;app.quit();}}}).finally(()=>{closing=false;});});
 current.on('closed',()=>{if(window===current)window=undefined;});
 void current.loadURL(server.url);
}
function setMenu(){const mac=process.platform==='darwin';const menu:Electron.MenuItemConstructorOptions[]=[
 ...(mac?[{label:'MyDoc',submenu:[{role:'about' as const},{type:'separator' as const},{role:'hide' as const},{role:'hideOthers' as const},{role:'unhide' as const},{type:'separator' as const},{label:'退出 MyDoc',accelerator:'Cmd+Q',click:()=>void requestQuit()}]}]:[]),
 {label:'文件',submenu:[{label:'打开文件…',accelerator:'CmdOrCtrl+O',click:()=>{focusWindow();if(window)void dialog.showOpenDialog(window,{title:'打开文件并导入',properties:['openFile','multiSelections'],filters:[{name:'Markdown / 文本 / ZIP',extensions:['md','markdown','txt','zip']}]}).then(result=>{if(!result.canceled)enqueueFiles(result.filePaths);});}},{label:'重新载入',accelerator:'CmdOrCtrl+R',click:()=>void reload()},{label:'打开数据目录',click:()=>{if(database)void shell.openPath(resolve(database.path,'..'));}},...(mac&&app.isPackaged?[{label:'卸载 MyDoc 并删除本机数据…',click:()=>void uninstallMac()}]:[]),{type:'separator'},mac?{role:'close'}:{label:'退出',accelerator:'Alt+F4',click:()=>void requestQuit()}]},
 {label:'编辑',submenu:[{role:'undo'},{role:'redo'},{type:'separator'},{role:'cut'},{role:'copy'},{role:'paste'},{role:'selectAll'}]},
 {label:'视图',submenu:[{role:'resetZoom'},{role:'zoomIn'},{role:'zoomOut'},{type:'separator'},{role:'togglefullscreen'}]},
 {label:'窗口',submenu:[{role:'minimize'},{role:'zoom'}]},
 {label:'帮助',submenu:[{label:'版本与更新…',click:()=>{focusWindow();window?.webContents.send('mydoc:show-updates');}}]},
 ];Menu.setApplicationMenu(Menu.buildFromTemplate(menu));}
if(!app.requestSingleInstanceLock())app.quit();else{
 for(const url of commandLineLinks(process.argv))documentLinks.enqueue(url);
 openFiles.enqueue(commandLineFiles(process.argv,process.cwd(),app.isPackaged));
 app.on('second-instance',(_event,argv,cwd)=>{for(const url of commandLineLinks(argv))enqueueLink(url);const paths=commandLineFiles(argv,cwd,app.isPackaged);if(paths.length)enqueueFiles(paths);else focusWindow();});
 app.on('before-quit',event=>{if(!allowingQuit){event.preventDefault();void requestQuit();}});
 app.on('window-all-closed',()=>{if(process.platform!=='darwin'&&!allowingQuit)void requestQuit();});
 app.on('activate',()=>{if(!window&&server)makeWindow();});
 app.on('will-quit',()=>{updater?.stop();void server?.close();database?.close();});
 void app.whenReady().then(async()=>{
  // Packaged applications register the protocol; development leaves installed ownership intact.
  if(app.isPackaged)app.setAsDefaultProtocolClient('mydoc');
  if(process.platform==='darwin')app.dock?.setIcon(appIcon);
  while(!database){try{const path=!app.isPackaged&&process.env.MYDOC_DESKTOP_DB_PATH?resolve(process.env.MYDOC_DESKTOP_DB_PATH):desktopDatabasePath(app.isPackaged,resolve(__dirname,'../../../../..'),app.getPath('userData'));if(app.isPackaged)migrateLegacyDesktopDatabase(dirname(app.getPath('exe')),app.getPath('userData'));database=new SQLiteProjectDatabase(path,connectDesktopDatabase,upgradeCloudDatabase,{backupExtension:LOCAL_OFFICE_BACKUP});}catch(error){const result=await dialog.showMessageBox({type:'error',title:'数据库打开失败',message:'无法打开本地数据库',detail:error instanceof Error?error.message:String(error),buttons:['重试','退出'],defaultId:0,cancelId:1});if(result.response===1){allowingQuit=true;app.quit();return;}}}
  try{server=await startDesktopServer(join(__dirname,'renderer'),database,electronCloudRequest,(url,options)=>net.fetch(url instanceof URL?url.href:url,options));updater=createDesktopUpdates({
   notify:state=>{if(state.phase==='installing'){allowingQuit=true;if(database){database.restoring=true;updateOwnsDatabaseLock=true;}}window?.webContents.send('mydoc:update-state',state);},
   release:()=>{if(!updateOwnsClose)return;updateOwnsClose=false;allowingQuit=false;closing=false;if(updateOwnsDatabaseLock&&database)database.restoring=false;updateOwnsDatabaseLock=false;window?.webContents.send('mydoc:update-unlock');},
   prepare:async()=>{if(closing||!server||!window)throw new Error('正在处理其他操作，请稍后安装更新。');updateOwnsClose=true;closing=true;if(!await saveBeforeClose(true))throw new Error('保存未完成，已保留当前窗口，请完成保存后重试。');await server.backups.create('before-update');},
  });
  setMenu();makeWindow();}catch(error){dialog.showErrorBox('桌面启动失败',error instanceof Error?error.message:String(error));allowingQuit=true;app.quit();}
 });
}
