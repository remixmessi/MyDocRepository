import {afterEach,expect,it,vi} from 'vitest';
import {checkGitHubRelease,githubLatestReleaseApi,githubRepository,parseGitHubRelease} from '../packages/common/nbstore/src/github-releases';
const asset=(name:string)=>({name,state:'uploaded',size:1024,browser_download_url:githubRepository+'/releases/download/v1.2.3/'+name});
const release=(names:string[])=>({tag_name:'v1.2.3',draft:false,prerelease:false,html_url:githubRepository+'/releases/tag/v1.2.3',body:'更新说明',published_at:'2026-10-14T00:00:00Z',assets:names.map(asset)});
afterEach(()=>vi.unstubAllGlobals());
it('从正式 Release 匹配系统和架构，优先选择安装包而非更新压缩包',()=>{
 const data=release(['MyDoc-1.2.3-mac-arm64.zip','MyDoc-1.2.3-mac-x64.dmg','MyDoc-1.2.3-mac-arm64.dmg','MyDoc-1.2.3-win-x64.exe']);
 expect(parseGitHubRelease(data,'darwin','arm64')).toMatchObject({version:'1.2.3',notes:'更新说明',installer:{name:'MyDoc-1.2.3-mac-arm64.dmg'}});
 expect(parseGitHubRelease(data,'darwin','x64').installer?.name).toBe('MyDoc-1.2.3-mac-x64.dmg');
 expect(parseGitHubRelease(data,'win32','x64').installer?.name).toBe('MyDoc-1.2.3-win-x64.exe');
 expect(parseGitHubRelease(data,'win32','arm64').installer).toBeUndefined();
});
it('无安装包仍返回版本；不将源代码、不明架构、其他系统和校验文件作为安装包',()=>{
 for(const name of ['source.zip','MyDoc-mac.dmg','MyDoc-win-arm64.exe','MyDoc-mac-x64-arm64.dmg','MyDoc-mac-arm64.dmg.blockmap','MyDoc-windows-mac-arm64.dmg'])expect(parseGitHubRelease(release([name]),'darwin','arm64').installer).toBeUndefined();
 expect(parseGitHubRelease(release(['MyDoc-mac-universal.dmg']),'darwin','arm64').installer).toBeDefined();
 expect(parseGitHubRelease(release(['MyDoc-linux-x86_64.AppImage']),'linux','x64').installer).toBeDefined();
});
it('拒绝草稿、预发布、非法版本和不属于仓库的链接',()=>{
 const data=release(['MyDoc-mac-arm64.dmg']);
 for(const patch of [{draft:true},{prerelease:true},{tag_name:'v1.2.3-beta.1'},{tag_name:'invalid'},{html_url:'https://evil.example/releases/tag/v1.2.3'}])expect(()=>parseGitHubRelease({...data,...patch},'darwin','arm64')).toThrow();
 for(const url of ['https://github.com.evil.example/remixmessi/MyDocRepository/releases/download/v1/a.dmg', 'https://github.com/other/repo/releases/download/v1/a.dmg','javascript:alert(1)']){
  expect(parseGitHubRelease({...data,assets:[{...data.assets[0],browser_download_url:url}]},'darwin','arm64').installer).toBeUndefined();
 }
});
it('只显示有效 SHA-256 和发布时间',()=>{
 const data=release(['MyDoc-mac-arm64.dmg']);
 expect(parseGitHubRelease({...data,assets:[{...data.assets[0],digest:'sha256:'+'a'.repeat(64)}]},'darwin','arm64').installer?.sha256).toBe('a'.repeat(64));
 expect(parseGitHubRelease({...data,published_at:'bad'},'darwin','arm64').publishedAt).toBeUndefined();
});
it('检查固定仓库，处理未发布、限流、服务错误和取消',async()=>{
 const fetchMock=vi.fn();vi.stubGlobal('fetch',fetchMock);const controller=new AbortController();
 fetchMock.mockResolvedValueOnce(new Response(JSON.stringify(release([])),{status:200}));
 expect((await checkGitHubRelease('darwin','arm64',controller.signal))?.version).toBe('1.2.3');
 expect(fetchMock).toHaveBeenCalledWith(githubLatestReleaseApi,expect.objectContaining({signal:controller.signal,credentials:'omit'}));
 fetchMock.mockResolvedValueOnce(new Response(null,{status:404}));expect(await checkGitHubRelease('darwin','arm64',controller.signal)).toBeNull();
 for(const status of [403,429]){fetchMock.mockResolvedValueOnce(new Response(null,{status}));const fallback=vi.fn(async()=>parseGitHubRelease(release([]),'darwin','arm64'));expect((await checkGitHubRelease('darwin','arm64',controller.signal,fallback))?.version).toBe('1.2.3');expect(fallback).toHaveBeenCalledOnce();}
 fetchMock.mockResolvedValueOnce(new Response(null,{status:403}));await expect(checkGitHubRelease('darwin','arm64',controller.signal)).rejects.toThrow('拒绝');
 fetchMock.mockResolvedValueOnce(new Response(null,{status:403,headers:{'x-ratelimit-remaining':'0'}}));await expect(checkGitHubRelease('darwin','arm64',controller.signal)).rejects.toThrow('额度');
 fetchMock.mockResolvedValueOnce(new Response(null,{status:500}));await expect(checkGitHubRelease('darwin','arm64',controller.signal)).rejects.toThrow('500');
 fetchMock.mockRejectedValueOnce(new DOMException('Aborted','AbortError'));await expect(checkGitHubRelease('darwin','arm64',controller.signal)).rejects.toThrow('Aborted');
});
