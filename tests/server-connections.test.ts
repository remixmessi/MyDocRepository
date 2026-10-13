import {beforeEach,it,expect,vi} from 'vitest';
const preferences=vi.hoisted(()=>new Map<string,string>());
vi.mock('@mydoc/nbstore',()=>({projectPreferences:{getItem:(key:string)=>preferences.get(key)}}));
vi.mock('../packages/common/nbstore/src/cloud',()=>({cloudServerUrl:()=>''}));
import {normalizeServer,savedServers} from '../packages/frontend/core/src/server-connections';
beforeEach(()=>preferences.clear());
it('保存多个服务器并合并旧地址，规范化后去重',()=>{
 preferences.set('cloud-servers',JSON.stringify([{url:'https://one.example/',name:'公司'},{url:'https://two.example/docs/',name:'个人'},{url:'https://one.example',name:'重复'}]));
 preferences.set('cloud-server-url','https://three.example/');
 expect(savedServers()).toEqual([{url:'https://one.example',name:'公司'},{url:'https://two.example/docs',name:'个人'},{url:'https://three.example',name:'three.example'}]);
});
it('旧的单服务器配置仍可使用，损坏列表不会阻止恢复',()=>{
 preferences.set('cloud-servers','invalid');preferences.set('cloud-server-url','http://localhost:4182/');
 expect(savedServers()).toEqual([{url:'http://localhost:4182',name:'localhost:4182'}]);
});
it('拒绝携带凭据、查询参数或非 HTTP 协议的地址',()=>{
 for(const url of ['file:///tmp/docs','https://user:secret@example.com','https://example.com?token=abc','https://example.com#docs'])expect(()=>normalizeServer(url)).toThrow();
 expect(normalizeServer(' https://EXAMPLE.com/docs/// ')).toBe('https://example.com/docs');
});
