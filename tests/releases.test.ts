import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync,writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {createHash} from 'node:crypto';
import Database from 'better-sqlite3';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {compareReleaseVersions,validReleaseVersion} from '../packages/common/nbstore/src/releases';

async function withServer(test:(root:string,cookie:string,cloud:CloudServer)=>Promise<void>,databasePath?:string){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-release-')),cloud=new CloudServer(databasePath??join(dir,'cloud.sqlite'),{defaultAdmin:{email:'admin@example.com',password:'abcdefghijk',name:'管理员'}}),server=createServer(cloud.handler);
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 try{const login=await fetch(root+'/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'admin@example.com',password:'abcdefghijk'})});expect(login.status).toBe(200);await test(root,login.headers.get('set-cookie')!.split(';')[0],cloud);}finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
}
function upload(root:string,cookie:string,version:string,arch?:string,platform='darwin',filename='文档库-'+(arch??'unknown')+'.dmg'){
 return fetch(root+'/admin/release',{method:'POST',headers:{Cookie:cookie,'X-MyDoc-Version':version,'X-MyDoc-Platform':platform,...(arch?{'X-MyDoc-Arch':arch}:{}),'X-MyDoc-Filename':encodeURIComponent(filename),'X-MyDoc-Notes':Buffer.from('修复同步问题').toString('base64')},body:Buffer.from('binary-release-'+arch)});
}

it('macOS x64 与 arm64 独立发布、查找与删除，下载保留字节及中文文件名',async()=>withServer(async(root,cookie)=>{
 const arm=await upload(root,cookie,'1.2.3','arm64');expect(arm.status).toBe(201);const release=await arm.json();expect(release.arch).toBe('arm64');expect(release.notes).toBe('修复同步问题');expect(release.sha256).toBe(createHash('sha256').update('binary-release-arm64').digest('hex'));expect(release.path).toBeUndefined();
 expect((await upload(root,cookie,'2.0.0','x64')).status).toBe(201);
 const latest=(arch:string)=>fetch(root+'/releases/latest?platform=darwin&arch='+arch).then(response=>response.json());
 expect((await latest('arm64')).version).toBe('1.2.3');expect((await latest('x64')).version).toBe('2.0.0');
 const downloaded=await fetch(root+release.downloadPath.replace('/api/cloud',''));expect(downloaded.status).toBe(200);expect(downloaded.headers.get('content-disposition')).toContain("filename*=UTF-8''%E6%96%87");expect(await downloaded.text()).toBe('binary-release-arm64');
 const list=await (await fetch(root+'/admin/releases',{headers:{Cookie:cookie}})).json();expect(list.map((item:{arch:string})=>item.arch).sort()).toEqual(['arm64','x64']);
 expect((await fetch(root+'/admin/release',{method:'DELETE',headers:{Cookie:cookie,'Content-Type':'application/json'},body:JSON.stringify({id:release.id})})).status).toBe(200);expect(await latest('arm64')).toBeNull();expect((await latest('x64')).version).toBe('2.0.0');
}));

it('拒绝缺失或无效 macOS 架构，上传及管理要求服务器管理员权限',async()=>withServer(async(root,cookie)=>{
 for(const arch of [undefined,'universal','amd64'])expect((await upload(root,cookie,'1.2.3',arch)).status).toBe(400);
 for(const query of ['', '&arch=universal'])expect((await fetch(root+'/releases/latest?platform=darwin'+query)).status).toBe(400);
 expect((await fetch(root+'/admin/releases')).status).toBe(401);expect((await upload(root,'','1.2.3','arm64')).status).toBe(401);
 const registration=await fetch(root+'/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'member@example.com',password:'abcdefghijk',name:'普通用户'})});const memberCookie=registration.headers.get('set-cookie')!.split(';')[0];expect((await upload(root,memberCookie,'1.2.3','arm64')).status).toBe(403);
 expect((await upload(root,cookie,'1.2.3',undefined,'win32')).status).toBe(201);expect((await (await fetch(root+'/releases/latest?platform=win32&arch=x64')).json()).version).toBe('1.2.3');
}));

it('最新版本按语义版本比较，较晚上传的旧版本不覆盖，同版稳定包优于预发布包',async()=>withServer(async(root,cookie)=>{
 for(const version of ['1.3.0','1.2.9','1.3.0-beta.2','1.3.0-beta.10'])expect((await upload(root,cookie,version,'arm64')).status).toBe(201);
 expect((await (await fetch(root+'/releases/latest?platform=darwin&arch=arm64')).json()).version).toBe('1.3.0');
}));

it('旧数据库补充架构列，未标注的 macOS 包不会匹配任一架构，重启保留新记录',async()=>{
 const directory=mkdtempSync(join(tmpdir(),'mydoc-release-migration-')),path=join(directory,'cloud.sqlite');
 try{
  await withServer(async(root,cookie)=>{expect((await upload(root,cookie,'9.0.0','arm64')).status).toBe(201);},path);
  const db=new Database(path);db.exec('ALTER TABLE cloud_releases DROP COLUMN arch');db.close();
  await withServer(async(root,cookie)=>{for(const arch of ['arm64','x64'])expect(await (await fetch(root+'/releases/latest?platform=darwin&arch='+arch)).json()).toBeNull();const items=await (await fetch(root+'/admin/releases',{headers:{Cookie:cookie}})).json();expect(items[0].arch).toBe('');expect((await upload(root,cookie,'1.0.0','x64')).status).toBe(201);},path);
  await withServer(async(root)=>{const release=await (await fetch(root+'/releases/latest?platform=darwin&arch=x64')).json();expect(release.version).toBe('1.0.0');expect(release.arch).toBe('x64');},path);
 }finally{rmSync(directory,{recursive:true,force:true});}
});

it('版本号验证与 SemVer 的预发布和构建标识优先级一致',()=>{
 expect(validReleaseVersion('1.2.3-beta.1+build.4')).toBe(true);for(const value of ['01.2.3','1.2','1.2.3-beta.01','1.2.3-'])expect(validReleaseVersion(value)).toBe(false);
 expect(compareReleaseVersions('1.0.0-beta.10','1.0.0-beta.2')).toBeGreaterThan(0);expect(compareReleaseVersions('1.0.0','1.0.0-rc.1')).toBeGreaterThan(0);expect(compareReleaseVersions('1.0.0+arm64','1.0.0+x64')).toBe(0);
});


it('服务器自动生成分架构更新清单，下载字节、范围与 HEAD 正确',async()=>withServer(async(root,cookie)=>{
 for(const [platform,arch,filename] of [['darwin','arm64','客户端-arm.zip'],['darwin','x64','客户端-intel.zip'],['win32','x64','客户端.exe']]){
  const uploaded=await upload(root,cookie,'2.0.0',arch,platform,filename);expect(uploaded.status).toBe(201);const row=await uploaded.json();
  const base=root+'/releases/update/'+platform+'/'+arch+'/',file=platform==='darwin'?'latest-mac.yml':'latest.yml';
  const response=await fetch(base+file);expect(response.status).toBe(200);const manifest=await response.json();
  expect(manifest).toMatchObject({version:'2.0.0',mydocPlatform:platform,mydocArch:arch,releaseNotes:'修复同步问题',files:[{url:row.id+(platform==='darwin'?'.zip':'.exe'),size:row.bytes,sha512:createHash('sha512').update('binary-release-'+arch).digest('base64')}]});
  const download=base+manifest.files[0].url;expect(await (await fetch(download)).text()).toBe('binary-release-'+arch);
  const range=await fetch(download,{headers:{Range:'bytes=2-6'}});expect(range.status).toBe(206);expect(await range.text()).toBe('nary-');
  const head=await fetch(download,{method:'HEAD'});expect(head.headers.get('content-length')).toBe(String(row.bytes));expect(await head.text()).toBe('');
  expect((await fetch(download,{headers:{Range:'bytes=999-'}})).status).toBe(416);
  expect((await fetch(download.replace('/'+arch+'/','/'+(arch==='x64'?'arm64':'x64')+'/'))).status).toBe(404);
 }
}));
it('仅提供最新正式版的 ZIP/EXE，缺包时不回退旧版或其他架构',async()=>withServer(async(root,cookie)=>{
 const base=root+'/releases/update/darwin/arm64/latest-mac.yml';
 expect((await fetch(base)).status).toBe(404);
 await upload(root,cookie,'1.0.0','arm64','darwin','old.zip');await upload(root,cookie,'2.0.0','arm64','darwin','new.dmg');
 expect((await fetch(base)).status).toBe(409);await upload(root,cookie,'2.0.0','x64','darwin','intel.zip');expect((await fetch(base)).status).toBe(409);
 await upload(root,cookie,'2.0.0','arm64','darwin','arm.zip');await upload(root,cookie,'3.0.0-beta.1','arm64','darwin','beta.zip');
 expect((await (await fetch(base)).json()).version).toBe('2.0.0');
 // Old Windows rows without architecture remain available to the old download API, not to in-app installation.
 await upload(root,cookie,'2.0.0',undefined,'win32','legacy.exe');expect((await fetch(root+'/releases/update/win32/x64/latest.yml')).status).toBe(404);
}));
it('生成更新清单前核对实际文件与已有 SHA-256，不为损坏的包生成校验值',async()=>withServer(async(root,cookie,cloud)=>{
 const row=await (await upload(root,cookie,'2.0.0','arm64','darwin','app.zip')).json();
 const stored=cloud.database.db.prepare('SELECT path FROM cloud_releases WHERE id=?').get(row.id) as {path:string};writeFileSync(stored.path,'corrupted-release');
 const response=await fetch(root+'/releases/update/darwin/arm64/latest-mac.yml');expect(response.status).toBe(409);expect((await response.json()).error).toContain('完整性');
}));
it('更新清单重启后重建，已有发布记录与安装文件保持不变',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-update-feed-restart-')),database=join(dir,'cloud.sqlite');let expected='';
 try{await withServer(async(root,cookie)=>{await upload(root,cookie,'2.0.0','arm64','darwin','app.zip');expected=await (await fetch(root+'/releases/update/darwin/arm64/latest-mac.yml')).text();},database);
  await withServer(async(root)=>{expect(await (await fetch(root+'/releases/update/darwin/arm64/latest-mac.yml')).text()).toBe(expected);},database);
 }finally{rmSync(dir,{recursive:true,force:true});}
});

it('生成的服务器清单可由 electron-updater 的 generic provider 实际解析与定位下载',async()=>withServer(async(root,cookie)=>{
 const row=await (await upload(root,cookie,'2.0.0','arm64','darwin','app.zip')).json();
 const {GenericProvider}=await import('electron-updater/out/providers/GenericProvider');
 const base=root+'/releases/update/darwin/arm64/';
 const executor={request:async(options:import('node:http').RequestOptions)=>{
  const url=(options.protocol??'http:')+'//'+options.hostname+(options.port?':'+options.port:'')+options.path;
  const response=await fetch(url);if(!response.ok)throw new Error('feed '+response.status);return response.text();
 }};
 const provider=new GenericProvider({provider:'generic',url:base},{isAddNoCacheQuery:false} as unknown as import('electron-updater').AppUpdater,{platform:'darwin',isUseMultipleRangeRequest:false,executor:executor as unknown as import('electron-updater/out/electronHttpExecutor').ElectronHttpExecutor});
 const info=await provider.getLatestVersion();expect(info.version).toBe('2.0.0');const resolved=provider.resolveFiles(info);expect(resolved).toHaveLength(1);expect(resolved[0].url.href).toBe(base+row.id+'.zip');expect(await (await fetch(resolved[0].url)).text()).toBe('binary-release-arm64');
}));
