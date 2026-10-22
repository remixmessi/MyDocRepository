import 'fake-indexeddb/auto';
import {describe,it,expect} from 'vitest';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import {IndexedDBStorage,IndexedDBAssetStorage} from '@mydoc/nbstore';
import {applyOfficeTextEdit} from '../packages/frontend/core/src/office-text';
import {LocalWorkspace} from '../packages/frontend/core/src/workspace';
import {validateOfficeTextUpdate,MAX_OFFICE_TEXT_LENGTH} from '../packages/frontend/apps/web/server/office-text';
import {exportBackup} from '../packages/frontend/core/src/backup';
import {parseDraft} from '../packages/frontend/core/src/draft';

describe('TXT 字符协作',()=>{
  it('两个离线客户端插入可合并，撤销只作用于本人',()=>{
    const a=new Y.Doc(),b=new Y.Doc();const ta=a.getText('text'),tb=b.getText('text');ta.insert(0,'起点');Y.applyUpdate(b,Y.encodeStateAsUpdate(a));
    const undo=new Y.UndoManager(ta,{trackedOrigins:new Set([LOCAL_ORIGIN])});
    applyOfficeTextEdit(ta,ta.toString(),'甲起点');applyOfficeTextEdit(tb,tb.toString(),'起点乙');
    const ua=Y.encodeStateAsUpdate(a),ub=Y.encodeStateAsUpdate(b);Y.applyUpdate(a,ub);Y.applyUpdate(b,ua);
    expect(ta.toString()).toBe('甲起点乙');expect(tb.toString()).toBe(ta.toString());undo.undo();expect(ta.toString()).toBe('起点乙');
    undo.destroy();a.destroy();b.destroy();
  });
  it('输入法组合期间的远端前缀不被覆盖，emoji 与空白保留',()=>{
    const live=new Y.Doc(),snapshot=new Y.Doc();const text=live.getText('text'),reference=snapshot.getText('text');text.insert(0,'A\t😀\n');Y.applyUpdate(snapshot,Y.encodeStateAsUpdate(live));const before=reference.toString();
    text.insert(0,'远端');applyOfficeTextEdit(text,before,'A中文\t😀\n',reference);expect(text.toString()).toBe('远端A中文\t😀\n');
    applyOfficeTextEdit(text,text.toString(),'远端A中文\t😃\n');expect(text.toString()).toBe('远端A中文\t😃\n');live.destroy();snapshot.destroy();
  });
  it('工作台复用单一文本实例，以原始资源 ID 保存并在关闭后恢复',async()=>{
    const storageName=crypto.randomUUID(),id='doc:'+crypto.randomUUID();
    const resource={id,kind:'text' as const,name:'测试.txt',mime:'text/plain',sourceFormat:'txt',originalAssetId:null,modelVersion:1,generation:'g',createdBy:'user',created:1,updated:1,deleted:0};
    const ws=new LocalWorkspace(new IndexedDBStorage(storageName),'test',true,new IndexedDBAssetStorage(crypto.randomUUID()));await ws.open();ws.officeResources.set(id,resource);
    const [first,second]=await Promise.all([ws.openOfficeText(id),ws.openOfficeText(id)]);expect(first).toBe(second);expect(first.persistence.id).toBe(id);
    applyOfficeTextEdit(first.text,'','<literal>\t中文😀\n');await ws.flushResource(id);const draft=await ws.exportDraft();expect(draft.office[0].text).toBe('<literal>\t中文😀\n');await expect(exportBackup(ws,'office')).rejects.toThrow('当前存储无法生成办公文件完整备份');expect(()=>parseDraft(JSON.stringify(draft))).toThrow('该草稿包含办公文件');await ws.destroy();
    const reopened=new LocalWorkspace(new IndexedDBStorage(storageName),'test',true,new IndexedDBAssetStorage(crypto.randomUUID()));await reopened.open();reopened.officeResources.set(id,resource);expect((await reopened.openOfficeText(id)).text.toString()).toBe('<literal>\t中文😀\n');await reopened.destroy();
  });
});

describe('TXT 服务端隔离校验',()=>{
  it('接受纯文本与幂等更新，拒绝其他根、格式、嵌入、伪造 map 和超限',()=>{
    const doc=new Y.Doc();doc.getText('text').insert(0,'hello');const valid=Y.encodeStateAsUpdate(doc);expect(()=>validateOfficeTextUpdate(null,valid)).not.toThrow();expect(()=>validateOfficeTextUpdate(valid,valid)).not.toThrow();
    const invalid=[()=>{const d=new Y.Doc();d.getText('body').insert(0,'污染');return d;},()=>{const d=new Y.Doc();d.getText('text').insert(0,'rich',{bold:true});return d;},()=>{const d=new Y.Doc();d.getText('text').insertEmbed(0,{image:'unsafe'});return d;},()=>{const d=new Y.Doc();d.getMap('text').set('notText','value');return d;},()=>{const d=new Y.Doc();d.getText('text').insert(0,'x'.repeat(MAX_OFFICE_TEXT_LENGTH+1));return d;}];
    for(const create of invalid){const bad=create();expect(()=>validateOfficeTextUpdate(valid,Y.encodeStateAsUpdate(bad))).toThrow();bad.destroy();}
    expect(()=>validateOfficeTextUpdate(valid,new Uint8Array([255]))).toThrow();doc.destroy();
  });
});

describe('TXT 字面查找与协作替换',()=>{
 it('按非重叠匹配定位，保留大小写、正则字符、多行、空白与 emoji 边界',async()=>{
  const {findTextRanges}=await import('../packages/frontend/core/src/office-text-find');
  expect(findTextRanges('aaa','aa').ranges).toEqual([{start:0,end:2}]);expect(findTextRanges('AaA','a').ranges).toEqual([{start:1,end:2}]);expect(findTextRanges('a.*b.*','.*').ranges).toEqual([{start:1,end:3},{start:4,end:6}]);expect(findTextRanges('😀\n\t😀\n','😀\n').ranges).toEqual([{start:0,end:3},{start:4,end:7}]);expect(findTextRanges('😀','\ud83d').ranges).toEqual([]);expect(findTextRanges('abc','').ranges).toEqual([]);
 });
 it('远端前插后匹配仍定位原字符，内部变化或删除后旧匹配拒绝且不改文档',async()=>{
  const {findTextRanges,captureTextMatch,resolveTextMatch,replaceTextMatches}=await import('../packages/frontend/core/src/office-text-find');const doc=new Y.Doc(),text=doc.getText('text');
  try{text.insert(0,'中文 中文');const match=captureTextMatch(text,findTextRanges(text.toString(),'中文').ranges[1]);text.insert(0,'远端');expect(resolveTextMatch(text,match)).toEqual({start:5,end:7});expect(replaceTextMatches(text,'中文','你好',match)).toBe(1);expect(text.toString()).toBe('远端中文 你好');expect(()=>replaceTextMatches(text,'中文','不应修改',match)).toThrow('已改变');expect(text.toString()).toBe('远端中文 你好');const next=captureTextMatch(text,{start:2,end:4});text.insert(3,'外部');expect(resolveTextMatch(text,next)).toBeUndefined();expect(()=>replaceTextMatches(text,'中文','x',next)).toThrow('已改变');}finally{doc.destroy();}
 });
 it('全部替换为单个本人可撤销事务，保留远端变化并支持空替换与字面美元符号',async()=>{
  const {replaceTextMatches}=await import('../packages/frontend/core/src/office-text-find');const a=new Y.Doc(),b=new Y.Doc(),text=a.getText('text');text.insert(0,'中文😀 中文😀\n');Y.applyUpdate(b,Y.encodeStateAsUpdate(a));const undo=new Y.UndoManager(text,{trackedOrigins:new Set([LOCAL_ORIGIN])});
  try{undo.stopCapturing();expect(replaceTextMatches(text,'中文','$&')).toBe(2);undo.stopCapturing();b.getText('text').insert(0,'远端');Y.applyUpdate(a,Y.encodeStateAsUpdate(b));expect(text.toString()).toBe('远端$&😀 $&😀\n');undo.undo();expect(text.toString()).toBe('远端中文😀 中文😀\n');undo.redo();expect(text.toString()).toBe('远端$&😀 $&😀\n');expect(replaceTextMatches(text,'$&','')).toBe(2);expect(text.toString()).toBe('远端😀 😀\n');}finally{undo.destroy();a.destroy();b.destroy();}
 });
 it('匹配配额、结果长度、无效代理字符或过期查询失败时不修改文本',async()=>{
  const {findTextRanges,captureTextMatch,replaceTextMatches,MAX_TEXT_MATCHES}=await import('../packages/frontend/core/src/office-text-find');const doc=new Y.Doc(),text=doc.getText('text');
  try{text.insert(0,'x'.repeat(MAX_TEXT_MATCHES+1));expect(findTextRanges(text.toString(),'x')).toMatchObject({truncated:true});const before=Y.encodeStateAsUpdate(doc);expect(()=>replaceTextMatches(text,'x','')).toThrow('1 万');expect(Y.encodeStateAsUpdate(doc)).toEqual(before);text.delete(0,text.length);text.insert(0,'甲甲');expect(()=>replaceTextMatches(text,'甲','x'.repeat(500001))).toThrow('100 万');expect(()=>replaceTextMatches(text,'甲','\ud83d')).toThrow('无效');const match=captureTextMatch(text,{start:0,end:1});expect(()=>replaceTextMatches(text,'乙','x',match)).toThrow('已改变');expect(text.toString()).toBe('甲甲');expect(replaceTextMatches(text,'甲','甲')).toBe(0);expect(replaceTextMatches(text,'不存在','')).toBe(0);}finally{doc.destroy();}
 });
});

describe('TXT 逻辑行号和字符位置',()=>{
 it('空文本、空行及末尾换行保留正确逻辑行',async()=>{
  const {textLineStarts,textLinePosition}=await import('../packages/frontend/core/src/office-text-lines');expect([...textLineStarts('')]).toEqual([0]);expect([...textLineStarts('a\n\n')]).toEqual([0,2,3]);const value='a\n\n';expect(textLinePosition(value,textLineStarts(value),1)).toEqual({line:1,column:2});expect(textLinePosition(value,textLineStarts(value),2)).toEqual({line:2,column:1});expect(textLinePosition(value,textLineStarts(value),3)).toEqual({line:3,column:1});
 });
 it('列和选中字符按码点计数，代理字符中间位置归到左边界',async()=>{
  const {textLineStarts,textLinePosition,textCharacterCount}=await import('../packages/frontend/core/src/office-text-lines');const value='甲😀\te\u0301\n末';const starts=textLineStarts(value);expect(textLinePosition(value,starts,3)).toEqual({line:1,column:3});expect(textLinePosition(value,starts,2)).toEqual({line:1,column:2});expect(textCharacterCount(value,1,3)).toBe(1);expect(textCharacterCount(value,0,6)).toBe(5);expect(textLinePosition(value,starts,100)).toEqual({line:2,column:2});expect(textLinePosition(value,starts,-1)).toEqual({line:1,column:1});
 });
 it('百万换行索引及可见窗口保持有界，不为每行创建可视节点',async()=>{
  const {textLineStarts,textLinePosition,visibleTextLines}=await import('../packages/frontend/core/src/office-text-lines');const value='\n'.repeat(1_000_000),starts=textLineStarts(value);expect(starts.length).toBe(1_000_001);expect(starts.byteLength).toBe(4_000_004);expect(textLinePosition(value,starts,value.length)).toEqual({line:1_000_001,column:1});const first=visibleTextLines(starts.length,0,320,25.5,24);expect(first.first).toBe(0);expect(first.last).toBeLessThan(20);const last=visibleTextLines(starts.length,25.5*1_000_000,320,25.5,24);expect(last.last).toBe(starts.length);expect(last.last-last.first).toBeLessThan(20);expect(visibleTextLines(1,0,0,0,0)).toEqual({first:0,last:1});
 });
});

it('TXT 在线选区随远端插入移动，忽略其他根和非法选区，限制成员列表',async()=>{
 const {captureTextCursor,textPeers}=await import('../packages/frontend/core/src/office-text-presence');const doc=new Y.Doc(),text=doc.getText('text');
 try{text.insert(0,'甲😀\n乙');const cursor=captureTextCursor(text,1,3,true);const states=new Map<number,Record<string,unknown>>([[1,{user:{name:'本人',color:'#000000'}}],[2,{user:{name:'成员',color:'url(x)'},cursor}]]);text.insert(0,'远端\n');const peers=textPeers(text,states,1);expect(peers[0]).toMatchObject({name:'成员',color:'#3159cb',anchor:6,head:4,startLine:2,endLine:2,label:'选区：第 2 行 2 列至第 2 行 3 列'});const wrong=doc.getText('other');wrong.insert(0,'其他根');states.set(3,{user:{name:'错误根'},cursor:captureTextCursor(wrong,0,1)});states.set(4,{user:{name:'坏位置'},cursor:{anchor:'x',head:null}});expect(textPeers(text,states,1).slice(1).map(peer=>peer.label)).toEqual(['正在查看','正在查看']);for(let i=5;i<50;i++)states.set(i,{user:{name:'在线'}});expect(textPeers(text,states,1)).toHaveLength(32);}finally{doc.destroy();}
});
