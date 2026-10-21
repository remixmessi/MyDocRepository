import {it,expect} from 'vitest';
import {createServer,type IncomingMessage} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {startDesktopServer} from '../packages/frontend/apps/desktop/src/server';
import {ProjectDatabase} from '../packages/frontend/apps/web/sqlite-server';
import {cloudConnectionError} from '../packages/frontend/apps/desktop/src/cloud-proxy';

it('连接错误保留错误码，macOS 路由或权限错误提示本地网络设置',()=>{
  expect(cloudConnectionError({name:'Error',message:'',code:'EHOSTUNREACH'},'http://192.168.0.85:19808','darwin')).toContain('隐私与安全 > 本地网络');
  expect(cloudConnectionError({name:'Error',message:'',code:'ECONNREFUSED'},'http://192.168.0.85:19808','win32')).toContain('ECONNREFUSED');
});

it('桌面代理隔离服务器 Cookie，转发二进制和实时流，支持会话轮换及退出',async()=>{
  const directory=mkdtempSync(join(tmpdir(),'mydoc-cloud-proxy-'));
  const db=new ProjectDatabase(join(directory,'db.sqlite'));
  let session='initial-session';
  const requests:IncomingMessage[]=[];
  const remote=createServer(async(req,res)=>{
    requests.push(req);
    const path=new URL(req.url!,'http://test').pathname;
    const loggedIn=req.headers.cookie==='mydoc_session='+session;
    if(path.endsWith('/login')||path.endsWith('/password')){
      if(path.endsWith('/password')&&!loggedIn){res.writeHead(401);res.end();return;}
      if(path.endsWith('/password'))session='rotated-session';
      res.setHeader('Set-Cookie','mydoc_session='+session+'; HttpOnly; SameSite=None; Secure; Path=/; Max-Age=604800');
      res.setHeader('Content-Type','application/json');res.end('{"ok":true}');return;
    }
    if(path.endsWith('/session')){res.setHeader('Content-Type','application/json');res.end(JSON.stringify({user:loggedIn?{name:'User'}:null}));return;}
    if(!loggedIn){res.writeHead(401);res.end();return;}
    if(path.endsWith('/logout')){session='logged-out';res.setHeader('Set-Cookie','mydoc_session=; HttpOnly; Path=/; Max-Age=0');res.end('{}');return;}
    if(path.endsWith('/asset')){
      const chunks:Buffer[]=[];for await(const chunk of req)chunks.push(Buffer.from(chunk));
      res.setHeader('Content-Type','application/octet-stream');res.end(Buffer.concat(chunks));return;
    }
    if(path.endsWith('/office-text-import')){
      const chunks:Buffer[]=[];for await(const chunk of req)chunks.push(Buffer.from(chunk));
      res.setHeader('Content-Type','application/json');res.end(JSON.stringify({bytes:[...Buffer.concat(chunks)]}));return;
    }
    if(path.endsWith('/office-export')){
      res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');res.setHeader('Content-Disposition',"attachment; filename*=UTF-8''copy.docx");res.setHeader('X-MyDoc-Snapshot','a'.repeat(64));res.end(Buffer.from([80,75,3,4]));return;
    }
    if(path.endsWith('/office-restore')||path.endsWith('/workspace-restore')){const chunks:Buffer[]=[];for await(const chunk of req)chunks.push(Buffer.from(chunk));res.setHeader('Content-Type','application/json');res.end(Buffer.concat(chunks));return;}
    if(path.endsWith('/workspace-backup')){res.setHeader('Content-Type','application/json');res.setHeader('Content-Disposition',"attachment; filename*=UTF-8''team.mydoc-team.json");res.end('{"type":"mydoc-team-workspace"}');return;}
    if(path.endsWith('/office-backup')){res.setHeader('Content-Type','application/json');res.setHeader('Content-Disposition',"attachment; filename*=UTF-8''file.mydoc-office.json");res.end('{"version":1}');return;}
    if(path.endsWith('/office-original')){
      res.setHeader('Content-Type','application/octet-stream');res.setHeader('Content-Disposition',"attachment; filename*=UTF-8''original.txt");res.end(Buffer.from([9,8,7]));return;
    }
    if(path.endsWith('/admin/release')){
      const chunks:Buffer[]=[];for await(const chunk of req)chunks.push(Buffer.from(chunk));
      res.setHeader('Content-Type','application/json');res.end(JSON.stringify({arch:req.headers['x-mydoc-arch'],version:req.headers['x-mydoc-version'],filename:req.headers['x-mydoc-filename'],bytes:[...Buffer.concat(chunks)]}));return;
    }
    if(['/admin/users','/admin/user','/admin/workspaces','/admin/workspace','/admin/workspace-member','/admin/workspace-owner','/admin/workspace-audit'].some(route=>path.endsWith(route))){res.setHeader('Content-Type','application/json');res.end('{"ok":true}');return;}
    if(path.endsWith('/admin/releases')||path.endsWith('/releases/latest')){res.setHeader('Content-Type','application/json');res.end('[]');return;}
    if(path.endsWith('/releases/download')){res.setHeader('Content-Disposition',"attachment; filename=MyDoc.dmg; filename*=UTF-8''MyDoc.dmg");res.end('installer');return;}
    if(path.endsWith('/events')){res.setHeader('Content-Type','text/event-stream');res.write('data: {"id":"doc:test"}\n\n');return;}
    res.writeHead(404);res.end();
  });
  await new Promise<void>(done=>remote.listen(0,'127.0.0.1',done));
  const desktop=await startDesktopServer(directory,db);
  const base='http://127.0.0.1:'+(remote.address() as {port:number}).port+'/docs';
  const prefix='/api/cloud-proxy/'+encodeURIComponent(base)+'/api/cloud';
  try{
    const bootstrap=await fetch(desktop.url,{redirect:'manual'});
    const desktopCookie=bootstrap.headers.get('set-cookie')!.split(';')[0];
    let cloudCookie='';
    const call=(path:string,init:RequestInit={})=>fetch(desktop.origin+prefix+path,{...init,headers:{...Object.fromEntries(new Headers(init.headers)),Cookie:desktopCookie+'; '+cloudCookie}});
    expect((await fetch(desktop.origin+prefix+'/login',{method:'POST'})).status).toBe(403);
    expect((await call('/login',{method:'POST',headers:{Origin:'http://evil.example'}})).status).toBe(403);
    const login=await call('/login',{method:'POST',headers:{Origin:desktop.origin},body:'{}'});
    expect(login.status).toBe(200);
    const stored=login.headers.get('set-cookie')!;
    expect(stored).toContain('HttpOnly');expect(stored).toContain('SameSite=Strict');expect(stored).not.toContain('Secure');
    expect(stored).toContain('Path=/api/cloud-proxy/'+encodeURIComponent(base)+'/');
    cloudCookie=stored.split(';')[0];
    expect((await (await call('/session')).json()).user.name).toBe('User');
    expect(requests.at(-1)!.url).toBe('/docs/api/cloud/session');
    expect(requests.at(-1)!.headers.cookie).toBe('mydoc_session='+session);
    expect(requests.at(-1)!.headers.origin).toBeUndefined();
    const otherPrefix='/api/cloud-proxy/'+encodeURIComponent(base+'/other')+'/api/cloud';
    const other=await fetch(desktop.origin+otherPrefix+'/session',{headers:{Cookie:desktopCookie+'; '+cloudCookie}});
    expect((await other.json()).user).toBeNull();
    const bytes=new Uint8Array([0,255,1,128]);
    const asset=await call('/asset?id=file',{method:'PUT',body:bytes,headers:{'Content-Type':'application/octet-stream'}});
    expect(new Uint8Array(await asset.arrayBuffer())).toEqual(bytes);
    const imported=await call('/office-text-import?workspace=test&name=原件.txt',{method:'POST',body:bytes,headers:{'Content-Type':'application/octet-stream'}});
    expect(await imported.json()).toEqual({bytes:[...bytes]});
    const original=await call('/office-original?workspace=test&id=doc:test');
    expect(original.headers.get('content-disposition')).toContain("filename*=UTF-8''original.txt");expect(new Uint8Array(await original.arrayBuffer())).toEqual(new Uint8Array([9,8,7]));
    const archive=await call('/office-backup?workspace=test&id=doc:test');expect(archive.headers.get('content-type')).toContain('application/json');expect(archive.headers.get('content-disposition')).toContain('mydoc-office.json');expect(await archive.json()).toEqual({version:1});const recovery=await call('/office-restore?workspace=test',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({type:'mydoc-office-resource'})});expect(await recovery.json()).toEqual({type:'mydoc-office-resource'});
    const full=await call('/workspace-backup?workspace=test');expect(full.headers.get('content-disposition')).toContain('mydoc-team.json');expect(await full.json()).toEqual({type:'mydoc-team-workspace'});const restored=await call('/workspace-restore?operationId=test',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({version:1})});expect(await restored.json()).toEqual({version:1});
    const exported=await call('/office-export?workspace=test&id=doc:test&format=docx&scope=paragraphs');
    expect(exported.status).toBe(200);expect(exported.headers.get('content-type')).toContain('wordprocessingml');expect(exported.headers.get('content-disposition')).toContain('copy.docx');expect(exported.headers.get('x-mydoc-snapshot')).toBe('a'.repeat(64));expect([...new Uint8Array(await exported.arrayBuffer())]).toEqual([80,75,3,4]);
    const publication=await call('/admin/release',{method:'POST',body:bytes,headers:{'X-MyDoc-Version':'1.2.3','X-MyDoc-Arch':'arm64','X-MyDoc-Platform':'darwin','X-MyDoc-Filename':'MyDoc-arm64.dmg'}});
    expect(await publication.json()).toEqual({arch:'arm64',version:'1.2.3',filename:'MyDoc-arm64.dmg',bytes:[...bytes]});
    expect((await call('/admin/releases')).status).toBe(200);expect((await call('/releases/latest?platform=darwin&arch=arm64')).status).toBe(200);
    expect(requests.at(-1)!.url).toBe('/docs/api/cloud/releases/latest?platform=darwin&arch=arm64');
    expect((await call('/admin/users?workspace=test')).status).toBe(200);expect((await call('/admin/user',{method:'PATCH',body:'{}'})).status).toBe(200);
    for(const route of ['/admin/workspaces','/admin/workspace','/admin/workspace-member','/admin/workspace-owner','/admin/workspace-audit'])expect((await call(route+'?workspace=test')).status).toBe(200);
    const installer=await call('/releases/download?id=release');expect(installer.headers.get('content-disposition')).toContain('filename*=UTF-8');expect(await installer.text()).toBe('installer');
    const events=await call('/events?workspace=test');
    expect(events.headers.get('content-type')).toBe('text/event-stream');
    const reader=events.body!.getReader();expect(new TextDecoder().decode((await reader.read()).value)).toContain('doc:test');await reader.cancel();
    const rotated=await call('/password',{method:'POST',body:'{}'});
    expect(rotated.status).toBe(200);cloudCookie=rotated.headers.get('set-cookie')!.split(';')[0];
    expect((await (await call('/session')).json()).user.name).toBe('User');
    const logout=await call('/logout',{method:'POST'});
    expect(logout.headers.get('set-cookie')).toContain('Max-Age=0');cloudCookie='';
    expect((await (await call('/session')).json()).user).toBeNull();
  }finally{await desktop.close();remote.closeAllConnections();await new Promise<void>(done=>remote.close(()=>done()));db.close();rmSync(directory,{recursive:true,force:true});}
});
