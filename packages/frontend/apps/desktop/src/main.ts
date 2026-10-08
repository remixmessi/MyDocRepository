import {app,BrowserWindow,Menu,ipcMain,dialog,shell} from 'electron';
import {join,resolve} from 'node:path';
import {randomUUID} from 'node:crypto';
import {mkdirSync} from 'node:fs';
import {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {connectDesktopDatabase} from './sqlite-driver';
import {startDesktopServer} from './server';
import {desktopDatabasePath,safeExternalUrl,validStorageRequest} from './policy';
import type {DesktopStorageReply} from '../../../../common/nbstore/src/desktop';
app.setName('MyDoc');
const profile=!app.isPackaged&&process.env.MYDOC_DESKTOP_PROFILE?resolve(process.env.MYDOC_DESKTOP_PROFILE):join(app.getPath('appData'),app.isPackaged?'MyDoc':'MyDoc Development');
mkdirSync(profile,{recursive:true});app.setPath('userData',profile);
let database:SQLiteProjectDatabase|undefined,server:Awaited<ReturnType<typeof startDesktopServer>>|undefined,window:BrowserWindow|undefined;
let allowingQuit=false,closing=false;
const pendingClose=new Map<string,{resolve:(result:{ok:boolean;error?:string})=>void;sender:number}>();
function trusted(event:Electron.IpcMainInvokeEvent|Electron.IpcMainEvent){return !!window&&event.sender===window.webContents&&event.senderFrame===event.sender.mainFrame&&!!server&&new URL(event.senderFrame.url).origin===server.origin;}
ipcMain.handle('mydoc:storage',async(event,input):Promise<DesktopStorageReply>=>{
 if(!trusted(event)||!validStorageRequest(input)||!server)throw new Error('桌面存储请求无效');
 const response=await fetch(server.origin+'/api/storage'+input.path,{method:input.method,headers:{Authorization:'Bearer '+server.token},body:input.body?.slice().buffer});
 return {status:response.status,contentType:response.headers.get('content-type')??'application/octet-stream',body:new Uint8Array(await response.arrayBuffer())};
});
ipcMain.on('mydoc:close-result',(event,result:unknown)=>{if(!trusted(event)||!result||typeof result!=='object')return;const r=result as {id?:unknown;ok?:unknown;error?:unknown};if(typeof r.id!=='string'||typeof r.ok!=='boolean')return;const pending=pendingClose.get(r.id);if(pending&&pending.sender===event.sender.id)pending.resolve({ok:r.ok,error:typeof r.error==='string'?r.error.slice(0,2000):undefined});});
async function saveBeforeClose(){
 const target=window;if(!target||target.isDestroyed())return true;
 if(target.webContents.isLoading()){await dialog.showMessageBox(target,{type:'info',message:'应用正在打开，请稍后再关闭。'});return false;}
 const id=randomUUID();let timer:ReturnType<typeof setTimeout>;
 const result=await new Promise<{ok:boolean;error?:string}>(resolve=>{pendingClose.set(id,{resolve,sender:target.webContents.id});timer=setTimeout(()=>resolve({ok:false,error:'等待保存超时，请返回文档检查保存状态后重试。'}),15000);target.webContents.send('mydoc:prepare-close',id);});
 clearTimeout(timer!);pendingClose.delete(id);
 if(!result.ok){await dialog.showMessageBox(target,{type:'error',title:'尚未完成保存',message:'已保留窗口，请完成保存后再关闭。',detail:result.error??'保存失败，请重试。',buttons:['返回文档']});return false;}
 return true;
}
async function requestQuit(){if(allowingQuit||closing)return;closing=true;try{if(await saveBeforeClose()){allowingQuit=true;app.quit();}}finally{closing=false;}}
async function reload(){if(!closing&&await saveBeforeClose())window?.webContents.reload();}
function makeWindow(){
 if(!server)return;
 window=new BrowserWindow({width:1320,height:900,minWidth:760,minHeight:540,show:false,title:'MyDoc',backgroundColor:'#f5f5f5',webPreferences:{preload:join(__dirname,'preload.cjs'),contextIsolation:true,nodeIntegration:false,sandbox:true,webSecurity:true}});
 const current=window;
 current.webContents.setWindowOpenHandler(({url})=>{
  // Local PDF previews may open in a separate, unprivileged viewer.
  if(url.startsWith('blob:'+server!.origin+'/'))return {action:'allow',overrideBrowserWindowOptions:{title:'MyDoc 文件预览',webPreferences:{preload:'',contextIsolation:true,nodeIntegration:false,sandbox:true,webSecurity:true}}};
  const external=safeExternalUrl(url);if(external)void shell.openExternal(external);return {action:'deny'};
 });
 current.webContents.on('did-create-window',viewer=>{viewer.webContents.on('will-navigate',(event,url)=>{if(!url.startsWith('blob:'+server!.origin+'/'))event.preventDefault();});viewer.webContents.setWindowOpenHandler(()=>({action:'deny'}));});
 current.webContents.on('will-navigate',(event,url)=>{if(new URL(url).origin!==server?.origin){event.preventDefault();const external=safeExternalUrl(url);if(external)void shell.openExternal(external);}});
 current.webContents.on('will-attach-webview',event=>event.preventDefault());
 current.webContents.on('before-input-event',(event,input)=>{if(input.type==='keyDown'&&((input.control||input.meta)&&input.key.toLowerCase()==='r'||input.key==='F5')){event.preventDefault();void reload();}});
 current.webContents.session.setPermissionRequestHandler((contents,permission,callback,details)=>callback(contents===current.webContents&&details.isMainFrame&&['clipboard-sanitized-write','clipboard-read','fullscreen'].includes(permission)));
 current.webContents.session.on('will-download',(_event,item)=>item.setSaveDialogOptions({title:'保存文件',defaultPath:join(app.getPath('downloads'),item.getFilename())}));
 current.once('ready-to-show',()=>current.show());
 current.on('close',event=>{if(allowingQuit)return;event.preventDefault();if(closing)return;closing=true;void saveBeforeClose().then(ok=>{if(ok){if(process.platform==='darwin')current.destroy();else{allowingQuit=true;app.quit();}}}).finally(()=>{closing=false;});});
 current.on('closed',()=>{if(window===current)window=undefined;});
 void current.loadURL(server.url);
}
function setMenu(){const mac=process.platform==='darwin';const menu:Electron.MenuItemConstructorOptions[]=[
 ...(mac?[{label:'MyDoc',submenu:[{role:'about' as const},{type:'separator' as const},{role:'hide' as const},{role:'hideOthers' as const},{role:'unhide' as const},{type:'separator' as const},{label:'退出 MyDoc',accelerator:'Cmd+Q',click:()=>void requestQuit()}]}]:[]),
 {label:'文件',submenu:[{label:'重新载入',accelerator:'CmdOrCtrl+R',click:()=>void reload()},{label:'打开数据目录',click:()=>{if(database)void shell.openPath(resolve(database.path,'..'));}},{type:'separator'},mac?{role:'close'}:{label:'退出',accelerator:'Alt+F4',click:()=>void requestQuit()}]},
 {label:'编辑',submenu:[{role:'undo'},{role:'redo'},{type:'separator'},{role:'cut'},{role:'copy'},{role:'paste'},{role:'selectAll'}]},
 {label:'视图',submenu:[{role:'resetZoom'},{role:'zoomIn'},{role:'zoomOut'},{type:'separator'},{role:'togglefullscreen'}]},
 {label:'窗口',submenu:[{role:'minimize'},{role:'zoom'}]},
 ];Menu.setApplicationMenu(Menu.buildFromTemplate(menu));}
if(!app.requestSingleInstanceLock())app.quit();else{
 app.on('second-instance',()=>{if(window){if(window.isMinimized())window.restore();window.show();window.focus();}else if(server)makeWindow();});
 app.on('before-quit',event=>{if(!allowingQuit){event.preventDefault();void requestQuit();}});
 app.on('window-all-closed',()=>{if(process.platform!=='darwin'&&!allowingQuit)void requestQuit();});
 app.on('activate',()=>{if(!window&&server)makeWindow();});
 app.on('will-quit',()=>{void server?.close();database?.close();});
 void app.whenReady().then(async()=>{
  while(!database){try{const path=!app.isPackaged&&process.env.MYDOC_DESKTOP_DB_PATH?resolve(process.env.MYDOC_DESKTOP_DB_PATH):desktopDatabasePath(app.isPackaged,resolve(__dirname,'../../../../..'),app.getPath('userData'));database=new SQLiteProjectDatabase(path,connectDesktopDatabase);}catch(error){const result=await dialog.showMessageBox({type:'error',title:'数据库打开失败',message:'无法打开本地数据库',detail:error instanceof Error?error.message:String(error),buttons:['重试','退出'],defaultId:0,cancelId:1});if(result.response===1){allowingQuit=true;app.quit();return;}}}
  try{server=await startDesktopServer(join(__dirname,'renderer'),database);setMenu();makeWindow();}catch(error){dialog.showErrorBox('桌面启动失败',error instanceof Error?error.message:String(error));allowingQuit=true;app.quit();}
 });
}
