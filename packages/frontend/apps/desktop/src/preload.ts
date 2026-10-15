import {contextBridge,ipcRenderer} from 'electron';
import type {DesktopBridge} from '../../../../common/nbstore/src/desktop';
let prepareClose:()=>Promise<void>=async()=>{};
let previousInert=false;
async function prepare(id:string,lock=false){if(lock){previousInert=document.documentElement.inert;document.activeElement instanceof HTMLElement&&document.activeElement.blur();document.documentElement.inert=true;}try{await prepareClose();ipcRenderer.send('mydoc:close-result',{id,ok:true});}catch(error){if(lock)document.documentElement.inert=previousInert;ipcRenderer.send('mydoc:close-result',{id,ok:false,error:error instanceof Error?error.message:'保存失败'});}}
ipcRenderer.on('mydoc:prepare-close',(_event,id:string)=>void prepare(id));
ipcRenderer.on('mydoc:prepare-update',(_event,id:string)=>void prepare(id,true));
ipcRenderer.on('mydoc:update-unlock',()=>{document.documentElement.inert=previousInert;});
const bridge:DesktopBridge={
 platform:process.platform,
 arch:process.arch,
 checkServerUpdates:server=>ipcRenderer.invoke('mydoc:update-server',server),
 updateStatus:()=>ipcRenderer.invoke('mydoc:update-status'),
 updateAction:action=>ipcRenderer.invoke('mydoc:update-action',action),
 setAutomaticUpdateChecks:value=>ipcRenderer.invoke('mydoc:update-settings',value),
 onUpdate(callback){const listener=(_event:Electron.IpcRendererEvent,state:import('../../../../common/nbstore/src/updates').DesktopUpdateState)=>callback(state);ipcRenderer.on('mydoc:update-state',listener);return()=>ipcRenderer.removeListener('mydoc:update-state',listener);},
 onShowUpdates(callback){const listener=()=>callback();ipcRenderer.on('mydoc:show-updates',listener);return()=>ipcRenderer.removeListener('mydoc:show-updates',listener);},
 chooseBackupDirectory:()=>ipcRenderer.invoke('mydoc:choose-backup-directory'),
 openBackupDirectory:()=>ipcRenderer.invoke('mydoc:open-backup-directory'),
 takeDocumentLink:()=>ipcRenderer.invoke('mydoc:take-document-link'),
 acknowledgeDocumentLink:id=>ipcRenderer.invoke('mydoc:ack-document-link',id),
 onDocumentLink(callback){const listener=()=>callback();ipcRenderer.on('mydoc:document-link-available',listener);return()=>ipcRenderer.removeListener('mydoc:document-link-available',listener);},
 takeOpenFiles:()=>ipcRenderer.invoke('mydoc:take-open-files'),
 acknowledgeOpenFiles:id=>ipcRenderer.invoke('mydoc:ack-open-files',id),
 onOpenFiles(callback){const listener=()=>callback();ipcRenderer.on('mydoc:open-files-available',listener);return()=>ipcRenderer.removeListener('mydoc:open-files-available',listener);},
 storage:request=>ipcRenderer.invoke('mydoc:storage',request),
 onPrepareClose(callback){prepareClose=callback;return()=>{if(prepareClose===callback)prepareClose=async()=>{};};},
};
contextBridge.exposeInMainWorld('mydocDesktop',bridge);
