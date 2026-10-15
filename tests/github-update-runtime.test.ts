import {EventEmitter} from 'node:events';
import {createHash} from 'node:crypto';
import {afterEach,beforeEach,expect,it,vi} from 'vitest';
const mocks=vi.hoisted(()=>({app:{isPackaged:true,getVersion:()=> '0.1.0',getPath:()=>'/nonexistent/mydoc-update-tests'},readFileSync:vi.fn(),existsSync:vi.fn(()=>false),writeFileSync:vi.fn(),renameSync:vi.fn()}));
const native=Object.assign(new EventEmitter(),{setFeedURL:vi.fn(),checkForUpdates:vi.fn(),downloadUpdate:vi.fn(),quitAndInstall:vi.fn(),autoDownload:true,autoInstallOnAppQuit:true,autoRunAppAfterInstall:false,allowDowngrade:true,allowPrerelease:true,logger:null});
vi.mock('electron',()=>({app:mocks.app}));
vi.mock('electron-updater',()=>({autoUpdater:native}));
vi.mock('node:fs',()=>({readFileSync:mocks.readFileSync,existsSync:mocks.existsSync,writeFileSync:mocks.writeFileSync,renameSync:mocks.renameSync}));
const {createDesktopUpdates,githubUpdateFeed}=await import('../packages/frontend/apps/desktop/src/update-runtime');
const {githubUpdateFiles}=await import('../packages/frontend/apps/desktop/src/github-update-info');
const checksum=createHash('sha512').update('package').digest('base64');
const file=(os:string,arch:string,ext:string)=>({url:`MyDoc-9.0.0-${os}-${arch}.${ext}`,sha512:checksum});
const info=()=>({version:'9.0.0',releaseDate:'2026-10-15',releaseNotes:'修复更新',files:[file('mac','arm64','zip'),file('mac','x64','zip'),file('win','x64','exe')]});
let runtime:ReturnType<typeof createDesktopUpdates>|undefined;
const originalResourcesPath=Object.getOwnPropertyDescriptor(process,'resourcesPath');
const originalPlatform=Object.getOwnPropertyDescriptor(process,'platform')!,originalArch=Object.getOwnPropertyDescriptor(process,'arch')!;
beforeEach(()=>{Object.defineProperty(process,'resourcesPath',{configurable:true,value:'/fake/mydoc/resources'});Object.defineProperty(process,'platform',{value:'darwin'});Object.defineProperty(process,'arch',{value:'arm64'});vi.clearAllMocks();mocks.app.isPackaged=true;mocks.existsSync.mockReturnValue(false);mocks.readFileSync.mockImplementation(()=>{throw new Error('missing');});native.checkForUpdates.mockImplementation(async()=>{native.emit('update-available',info());return {};});native.downloadUpdate.mockImplementation(async()=>{native.emit('download-progress',{percent:50,transferred:50,total:100});native.emit('update-downloaded');return [];});});
afterEach(()=>{runtime?.stop();runtime=undefined;Object.defineProperty(process,'platform',originalPlatform);Object.defineProperty(process,'arch',originalArch);if(originalResourcesPath)Object.defineProperty(process,'resourcesPath',originalResourcesPath);else Reflect.deleteProperty(process,'resourcesPath');});
function setup(){const prepare=vi.fn(async()=>{}),release=vi.fn(),notify=vi.fn();runtime=createDesktopUpdates({prepare,release,notify});return {updates:runtime.updates,prepare,release,notify};}
it('packaged client uses fixed GitHub feed and downloads in-process before backup and install',async()=>{
 const {updates,prepare,notify}=setup();expect(native.setFeedURL).toHaveBeenCalledWith(githubUpdateFeed);expect(native.autoDownload).toBe(false);expect(native.autoInstallOnAppQuit).toBe(false);expect(native.allowDowngrade).toBe(false);expect(native.allowPrerelease).toBe(false);
 await runtime!.checkGitHub();expect(updates.status).toMatchObject({phase:'available',latestVersion:'9.0.0',source:'github',releaseNotes:'修复更新'});expect(native.downloadUpdate).not.toHaveBeenCalled();
 await updates.download();expect(updates.status.phase).toBe('downloaded');expect(notify.mock.calls.some(([state])=>state.percent===50)).toBe(true);expect(native.quitAndInstall).not.toHaveBeenCalled();
 prepare.mockImplementationOnce(async()=>{expect(native.quitAndInstall).not.toHaveBeenCalled();});await updates.install();expect(prepare).toHaveBeenCalledOnce();expect(native.quitAndInstall).toHaveBeenCalledWith(false,true);
});
it('download failure retries, backup failure prevents exit, asynchronous install failure unlocks',async()=>{
 const {updates,prepare,release}=setup();await updates.check();native.downloadUpdate.mockRejectedValueOnce(new Error('network'));await updates.download();expect(updates.status.retry).toBe('download');await updates.download();
 prepare.mockRejectedValueOnce(new Error('保护备份失败'));await updates.install();expect(native.quitAndInstall).not.toHaveBeenCalled();expect(updates.status).toMatchObject({phase:'downloaded',retry:'install'});expect(release).toHaveBeenCalledOnce();
 await updates.install();native.emit('error',new Error('install failed'));expect(release).toHaveBeenCalledTimes(2);expect(updates.status.phase).toBe('downloaded');
});
it('does not switch channels or restart checks while downloading or retaining a downloaded package',async()=>{
 const {updates}=setup();await updates.check();let finish!:()=>void;native.downloadUpdate.mockImplementationOnce(()=>new Promise<void>(resolve=>{finish=resolve;}));const downloading=updates.download();await runtime!.checkGitHub();expect(updates.selectChannel({source:'configured',enabled:true,channel:'other'})).toBe(false);expect(native.checkForUpdates).toHaveBeenCalledTimes(1);finish();await downloading;await runtime!.checkGitHub();expect(native.checkForUpdates).toHaveBeenCalledTimes(1);
});
it('development runtime never calls native check, download or install',async()=>{
 mocks.app.isPackaged=false;const {updates}=setup();await runtime!.checkGitHub();await updates.download();await updates.install();expect(updates.status.phase).toBe('disabled');expect(native.checkForUpdates).not.toHaveBeenCalled();expect(native.downloadUpdate).not.toHaveBeenCalled();expect(native.quitAndInstall).not.toHaveBeenCalled();
});
it('missing metadata gives an actionable release error and can be checked again',async()=>{
 const {updates}=setup();native.checkForUpdates.mockRejectedValueOnce(Object.assign(new Error('provider internals'),{code:'ERR_UPDATER_CHANNEL_FILE_NOT_FOUND'}));await updates.check();expect(updates.status.message).toContain('latest-mac.yml');expect(updates.status.retry).toBe('check');await updates.check();expect(updates.status.phase).toBe('available');
});
it('filters update packages by OS and architecture and rejects absent/invalid checksums',()=>{
 for(const [os,arch,ext,platform] of [['mac','arm64','zip','darwin'],['mac','x64','zip','darwin'],['win','x64','exe','win32']])expect(githubUpdateFiles(info(),platform,arch)).toEqual([file(os,arch,ext)]);
 for(const files of [[file('mac','x64','zip')],[file('mac','arm64','dmg')],[{...file('mac','arm64','zip'),sha512:'bad'}],[{...file('mac','arm64','zip'),url:'https://evil.example/mac-arm64.zip'}]])expect(()=>githubUpdateFiles({version:'9.0.0',files},'darwin','arm64')).toThrow();
});

it('a configured HTTPS channel switches to the fixed GitHub feed only while idle',async()=>{
 mocks.readFileSync.mockImplementation((path:unknown)=>String(path).endsWith('release-config.json')?JSON.stringify({enabled:true,channel:'公司渠道'}):JSON.stringify({automaticChecks:false}));mocks.existsSync.mockReturnValue(true);
 const {updates}=setup();expect(updates.status).toMatchObject({source:'configured',channel:'公司渠道',automaticChecks:false});expect(native.setFeedURL).not.toHaveBeenCalled();
 await runtime!.checkGitHub();expect(native.setFeedURL).toHaveBeenCalledWith(githubUpdateFeed);expect(updates.status).toMatchObject({source:'github',phase:'available',automaticChecks:false});
});
