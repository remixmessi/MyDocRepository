import {EventEmitter} from 'node:events';
import {createHash} from 'node:crypto';
import {afterEach,beforeEach,expect,it,vi} from 'vitest';
const mocks=vi.hoisted(()=>({app:{isPackaged:true,getVersion:()=> '0.1.0',getPath:()=>'/nonexistent/mydoc-update-tests'},readFileSync:vi.fn(),existsSync:vi.fn(()=>false),writeFileSync:vi.fn(),renameSync:vi.fn()}));
const native=Object.assign(new EventEmitter(),{setFeedURL:vi.fn(),checkForUpdates:vi.fn(),downloadUpdate:vi.fn(),quitAndInstall:vi.fn(),disableDifferentialDownload:false,autoDownload:true,autoInstallOnAppQuit:true,autoRunAppAfterInstall:false,allowDowngrade:true,allowPrerelease:true,logger:null});
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

const serverInfo=()=>({version:'9.0.0',mydocPlatform:'darwin',mydocArch:'arm64',releaseDate:'2026-10-15',releaseNotes:'服务器修复',files:[{url:'12345678-1234-1234-1234-123456789abc.zip',sha512:checksum,size:7}]});
it('server feed uses normalized address and real architecture, then reuses native download/backup/install',async()=>{
 const {updates,prepare}=setup();native.checkForUpdates.mockImplementationOnce(async()=>{native.emit('update-available',serverInfo());return {};});
 await runtime!.checkServer('  docs.example.com/mydoc/  ');
 expect(native.setFeedURL).toHaveBeenLastCalledWith({provider:'generic',url:'http://docs.example.com/mydoc/api/cloud/releases/update/darwin/arm64/',useMultipleRangeRequest:false});expect(native.disableDifferentialDownload).toBe(true);
 expect(updates.status).toMatchObject({source:'server',serverUrl:'http://docs.example.com/mydoc',phase:'available',releaseNotes:'服务器修复'});
 expect(native.downloadUpdate).not.toHaveBeenCalled();await updates.download();expect(updates.status.phase).toBe('downloaded');await updates.install();expect(prepare).toHaveBeenCalledOnce();expect(native.quitAndInstall).toHaveBeenCalledWith(false,true);
});
it('server selection and automatic check preference survive restart',async()=>{
 const {updates}=setup();native.checkForUpdates.mockImplementationOnce(async()=>{native.emit('update-available',serverInfo());return {};});await runtime!.checkServer('https://docs.example.com');updates.configure(false);
 const settings=String(mocks.writeFileSync.mock.calls.at(-1)![1]);expect(JSON.parse(settings)).toMatchObject({source:'server',serverUrl:'https://docs.example.com',automaticChecks:false});
 runtime!.stop();mocks.readFileSync.mockImplementation((path:unknown)=>{if(String(path).endsWith('update-settings.json'))return settings;throw new Error('missing');});
 const restarted=setup();expect(restarted.updates.status).toMatchObject({source:'server',serverUrl:'https://docs.example.com',automaticChecks:false});expect(native.setFeedURL).toHaveBeenLastCalledWith(expect.objectContaining({provider:'generic',url:'https://docs.example.com/api/cloud/releases/update/darwin/arm64/'}));
});
it('old server authentication failures explain the required deployment instead of reporting a network failure',async()=>{
 const {updates}=setup();native.checkForUpdates.mockRejectedValueOnce(Object.assign(new Error('sensitive provider response'),{statusCode:401,code:'HTTP_ERROR_401'}));
 await runtime!.checkServer('http://192.168.0.85:19808');
 expect(updates.status).toMatchObject({phase:'error',retry:'check',source:'server'});
 expect(updates.status.message).toContain('升级服务器');expect(updates.status.message).toContain('HTTP 401');expect(updates.status.message).not.toContain('sensitive');
});
it('server error messages distinguish HTTP failures, malformed manifests and connection failures without exposing provider details',async()=>{
 const {serverUpdateError}=await import('../packages/frontend/apps/desktop/src/update-runtime');
 for(const [error,expected] of [[{code:'HTTP_ERROR_403'},'匿名访问'],[{statusCode:409},'尚未就绪'],[{statusCode:502},'HTTP 502'],[{code:'ERR_UPDATER_CHANNEL_FILE_NOT_FOUND'},'升级服务器'],[{code:'ERR_UPDATER_INVALID_UPDATE_INFO'},'清单无效'],[{code:'ECONNREFUSED'},'地址、端口'],[{code:'ERR_CHECKSUM_MISMATCH'},'已阻止安装']] as const)expect(serverUpdateError(error)).toContain(expected);
 expect(serverUpdateError(new Error('secret URL'))).toBeUndefined();
});
it('server cannot replace a channel while downloading or holding a downloaded update',async()=>{
 const {updates}=setup();await updates.check();let finish!:()=>void;native.downloadUpdate.mockImplementationOnce(()=>new Promise<void>(resolve=>{finish=resolve;}));const download=updates.download();const calls=native.setFeedURL.mock.calls.length;
 await runtime!.checkServer('https://other.example.com');expect(updates.status.source).toBe('github');expect(native.setFeedURL).toHaveBeenCalledTimes(calls);finish();await download;await runtime!.checkServer('https://other.example.com');expect(updates.status.source).toBe('github');
});
it('server input rejects credentials and non-HTTP protocols, and metadata rejects wrong architecture',async()=>{
 const {updates}=setup();for(const url of ['file:///tmp/app.zip','https://user:secret@example.com','https://example.com/?token=x','https://example.com/#x',{},''])await expect(runtime!.checkServer(url)).rejects.toThrow();
 native.checkForUpdates.mockImplementationOnce(async()=>{native.emit('update-available',{...serverInfo(),mydocArch:'x64'});return {};});await runtime!.checkServer('http://127.0.0.1:8080');expect(updates.status).toMatchObject({source:'server',phase:'error',retry:'check'});expect(updates.status.message).toContain('架构');expect(native.downloadUpdate).not.toHaveBeenCalled();
});
it('switching back to GitHub resets server differential policy and persists the chosen channel',async()=>{
 setup();native.checkForUpdates.mockImplementationOnce(async()=>{native.emit('update-available',serverInfo());return {};});await runtime!.checkServer('http://127.0.0.1:8080');await runtime!.checkGitHub();expect(native.disableDifferentialDownload).toBe(false);expect(runtime!.updates.status.source).toBe('github');expect(JSON.parse(String(mocks.writeFileSync.mock.calls.at(-1)![1]))).toMatchObject({source:'github'});
});

it('server manifest contract rejects absolute URLs, prereleases, malformed checksums and wrong package formats',async()=>{
 const {serverUpdateFiles,serverUpdateFeed}=await import('../packages/common/nbstore/src/server-updates');
 const valid=serverInfo();expect(serverUpdateFiles(valid,'darwin','arm64')).toEqual(valid.files);
 for(const patch of [{version:'9.0.0-beta.1'},{mydocPlatform:'win32'},{files:[{...valid.files[0],url:'https://other.example/app.zip'}]},{files:[{...valid.files[0],url:valid.files[0].url.replace('.zip','.exe')}]},{files:[{...valid.files[0],sha512:'bad'}]}])expect(()=>serverUpdateFiles({...valid,...patch},'darwin','arm64')).toThrow();
 expect(serverUpdateFeed('https://docs.example.com/base/','win32','x64')).toBe('https://docs.example.com/base/api/cloud/releases/update/win32/x64/');
 expect(()=>serverUpdateFeed('https://docs.example.com','linux','x64')).toThrow();
});
