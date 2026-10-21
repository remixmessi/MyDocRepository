import {it,expect,vi} from 'vitest';
import {cloudRequest,configureCloudServer,registerLocalOffice,CloudAssets} from '../packages/common/nbstore/src/cloud';
it('个人办公请求及图片始终走当前 SQLite 服务，不发送到已配置的团队服务器',async()=>{
 const fetcher=vi.fn(async()=>new Response(JSON.stringify({id:'image',name:'图.png',mime:'image/png',size:3,data:'AQID'}),{headers:{'Content-Type':'application/json'}}));vi.stubGlobal('window',{});vi.stubGlobal('fetch',fetcher);configureCloudServer('https://team.example');const unregister=registerLocalOffice('local');
 try{for(const path of ['/office-resources','/office-import','/office-export','/office-original','/office-sheet-operation','/office-pdf-operation','/document'])await cloudRequest(path+'?workspace=local&id=doc:test');const assets=new CloudAssets('local');expect((await assets.getForDocument('local','image','doc:test'))?.bytes).toEqual(new Uint8Array([1,2,3]));expect(fetcher).toHaveBeenCalledTimes(8);for(const call of fetcher.mock.calls)expect(String((call as unknown[])[0])).toMatch(/^\/api\/storage\/office\//);
 }finally{unregister();configureCloudServer('');vi.unstubAllGlobals();}
});
it('桌面个人办公通过 IPC 传输原始二进制，不依赖 HTTP 登录或团队配置',async()=>{
 const storage=vi.fn(async(_request:unknown)=>({status:200,contentType:'application/octet-stream',body:new Uint8Array([80,75,1,2])})),fetcher=vi.fn();vi.stubGlobal('window',{mydocDesktop:{storage}});vi.stubGlobal('fetch',fetcher);configureCloudServer('https://team.example');const unregister=registerLocalOffice('local');
 try{const response=await cloudRequest('/office-original?workspace=local&id=doc:test');expect(new Uint8Array(await response.arrayBuffer())).toEqual(new Uint8Array([80,75,1,2]));expect(storage.mock.calls[0][0]).toMatchObject({path:'/office/office-original?workspace=local&id=doc:test',method:'GET'});expect(fetcher).not.toHaveBeenCalled();}
 finally{unregister();configureCloudServer('');vi.unstubAllGlobals();}
});

it('个人资源响应绑定代际，桌面 IPC 保留代际头，旧资源响应拒绝应用',async()=>{
 const {clearOfficeGenerations,officeDocumentGeneration,officeGenerationBlocked,cloudJson}=await import('../packages/common/nbstore/src/cloud');let generation='original-12345678';const storage=vi.fn(async()=>({status:200,contentType:'application/octet-stream',officeGeneration:generation,body:new Uint8Array([0,0])}));vi.stubGlobal('window',{mydocDesktop:{storage}});const unregister=registerLocalOffice('epoch-test');try{
  await cloudRequest('/document?workspace=epoch-test&id=doc:abc');expect(officeDocumentGeneration('epoch-test','doc:abc')).toBe(generation);generation='epoch-restored-12345';await expect(cloudRequest('/document?workspace=epoch-test&id=doc:abc')).rejects.toThrow('旧代际草稿');expect(officeGenerationBlocked('epoch-test','doc:abc')).toBe(true);
 }finally{clearOfficeGenerations('epoch-test');unregister();vi.unstubAllGlobals();}
});
