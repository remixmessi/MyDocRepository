import {contextBridge,ipcRenderer} from 'electron';
import type {DesktopBridge} from '../../../../common/nbstore/src/desktop';
let prepareClose:()=>Promise<void>=async()=>{};
ipcRenderer.on('mydoc:prepare-close',async(_event,id:string)=>{try{await prepareClose();ipcRenderer.send('mydoc:close-result',{id,ok:true});}catch(error){ipcRenderer.send('mydoc:close-result',{id,ok:false,error:error instanceof Error?error.message:'保存失败'});}});
const bridge:DesktopBridge={
 platform:process.platform,
 storage:request=>ipcRenderer.invoke('mydoc:storage',request),
 onPrepareClose(callback){prepareClose=callback;return()=>{if(prepareClose===callback)prepareClose=async()=>{};};},
};
contextBridge.exposeInMainWorld('mydocDesktop',bridge);
