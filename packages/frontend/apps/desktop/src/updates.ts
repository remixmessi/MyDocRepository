import type {DesktopUpdateState} from '../../../../common/nbstore/src/updates';
export interface UpdateDriver {check():Promise<unknown>;download():Promise<unknown>;install():void;}
interface Options {version:string;source?:DesktopUpdateState['source'];enabled:boolean;channel:string;disabledReason?:string;automaticChecks:boolean;driver?:UpdateDriver;prepare:()=>Promise<void>;release:()=>void;notify:(state:DesktopUpdateState)=>void;saveSettings:(automaticChecks:boolean)=>void;errorMessage?:(error:unknown)=>string|undefined;}
const busy=new Set(['checking','downloading','preparing','installing']);
export class DesktopUpdates {
 private state:DesktopUpdateState;
 constructor(private options:Options){this.state={version:options.version,source:options.source,enabled:options.enabled,automaticChecks:options.automaticChecks,channel:options.channel,phase:options.enabled?'idle':'disabled',message:options.disabledReason};}
 get status(){return {...this.state};}
 get canChangeChannel(){return !busy.has(this.state.phase)&&this.state.phase!=='downloaded';}
 selectChannel(value:{source:DesktopUpdateState['source'];channel:string;enabled:boolean;disabledReason?:string}){if(!this.canChangeChannel)return false;this.publish({source:value.source,channel:value.channel,enabled:value.enabled,phase:value.enabled?'idle':'disabled',message:value.disabledReason,latestVersion:undefined,releaseDate:undefined,releaseNotes:undefined,checkedAt:undefined,percent:undefined,transferred:undefined,total:undefined,retry:undefined});return true;}
 get installing(){return this.state.phase==='preparing'||this.state.phase==='installing';}
 private publish(patch:Partial<DesktopUpdateState>){this.state={...this.state,...patch};this.options.notify(this.status);}
 configure(value:unknown){if(typeof value!=='boolean')throw new Error('更新设置无效');this.options.saveSettings(value);this.publish({automaticChecks:value});return this.status;}
 async check(){if(!this.state.enabled||busy.has(this.state.phase)||this.state.phase==='downloaded')return this.status;this.publish({phase:'checking',message:undefined,retry:undefined,percent:undefined,latestVersion:undefined,releaseDate:undefined,releaseNotes:undefined,transferred:undefined,total:undefined});try{await this.options.driver!.check();if(this.state.phase==='checking')this.unavailable();}catch(error){this.failed('check',error);}return this.status;}
 available(version:string,date?:string,notes?:string){if(this.state.phase!=='checking')return;this.publish({phase:'available',latestVersion:version.slice(0,100),releaseDate:date,releaseNotes:notes?.slice(0,20000),checkedAt:Date.now(),message:undefined});}
 unavailable(){if(this.state.phase!=='checking')return;this.publish({phase:'up-to-date',latestVersion:undefined,checkedAt:Date.now(),message:undefined});}
 progress(percent:number,transferred:number,total:number){if(this.state.phase!=='downloading')return;this.publish({percent:Number.isFinite(percent)?Math.max(0,Math.min(100,percent)):0,transferred:Number.isFinite(transferred)?Math.max(0,transferred):0,total:Number.isFinite(total)?Math.max(0,total):0});}
 async download(){if(!this.state.enabled||!this.state.latestVersion||!['available','error'].includes(this.state.phase))return this.status;this.publish({phase:'downloading',percent:0,message:undefined,retry:undefined});try{await this.options.driver!.download();if(this.state.phase==='downloading')this.publish({phase:'downloaded',percent:100});}catch(error){this.failed('download',error);}return this.status;}
 downloaded(){if(this.state.phase==='downloading')this.publish({phase:'downloaded',percent:100,message:undefined});}
 failed(action?:'check'|'download'|'install',error?:unknown){
  const retry=action??(this.installing?'install':this.state.phase==='downloading'?'download':'check');
  if(this.installing)this.options.release();this.publish({phase:retry==='install'?'downloaded':'error',retry,message:this.options.errorMessage?.(error)??(retry==='install'?'保存、保护备份或安装未完成，已保留窗口与下载，请重试。':retry==='download'?'更新下载失败，请检查网络后重试。':'更新检查失败，请检查网络或稍后重试。')});
 }
 async install(){if(this.state.phase!=='downloaded')return this.status;this.publish({phase:'preparing',message:undefined,retry:undefined});try{await this.options.prepare();if(this.status.phase!=='preparing')return this.status;this.publish({phase:'installing'});this.options.driver!.install();}catch(error){this.failed('install');if(error instanceof Error)this.publish({message:error.message.slice(0,500)});}return this.status;}
}
