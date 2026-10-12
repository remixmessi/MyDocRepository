import 'fake-indexeddb/auto';
import {it,expect} from 'vitest';
import {Y} from '@mydoc/crdt';
import {IndexedDBStorage,IndexedDBAssetStorage} from '@mydoc/nbstore';
import {readDocument,writeMarkdown,SurfaceStore} from '@mydoc/editor-blocks';
import {LocalWorkspace} from '../packages/frontend/core/src/workspace';
import {CloudStorage} from '../packages/common/nbstore/src/cloud';
import {parseDraft,restoreDraft} from '../packages/frontend/core/src/draft';

it('工作空间草稿导出后可解析预览，恢复为新 ID 文档并保留正文与白板',async()=>{
 const ws=new LocalWorkspace(new IndexedDBStorage(crypto.randomUUID()),'local',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 await ws.open();
 const first=await ws.create('草稿文档甲','未提交的正文内容');
 const second=await ws.create('草稿文档乙','第二篇正文');
 new SurfaceStore((await ws.openDocument(first)).store).put({id:'shape',kind:'rectangle',x:0,y:0,width:50,height:50,fill:'#ffffff',stroke:'#000000',strokeWidth:2,text:'白板文字'});
 const draft=await ws.exportDraft();
 expect(draft.format).toBe('mydoc-workspace-draft');

 const preview=parseDraft(JSON.stringify(draft));
 expect(preview.kind).toBe('workspace');
 expect(preview.sourceWorkspaceId).toBe('local');
 expect(preview.documents).toHaveLength(2);
 const previewFirst=preview.documents.find(entry=>entry.sourceId===first)!;
 expect(previewFirst.title).toBe('草稿文档甲');
 expect(previewFirst.summary).toContain('未提交的正文内容');
 expect(previewFirst.surfaceElements).toBe(1);
 expect(previewFirst.missingAssets).toEqual([]);

 // 恢复到另一个个人工作区：必须是新 ID 的新文档，不复用草稿中的原 ID。
 const target=new LocalWorkspace(new IndexedDBStorage(crypto.randomUUID()),'local',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 await target.open();
 const folder=target.createFolder('恢复');
 const results=await restoreDraft(target,preview.documents,folder);
 expect(results).toHaveLength(2);
 expect(results.map(result=>result.id)).not.toContain(first);
 expect(results.map(result=>result.id)).not.toContain(second);
 const restored=target.list().find(meta=>meta.title==='草稿文档甲')!;
 expect(restored.folderId).toBe(folder);
 expect(readDocument((await target.openDocument(restored.id)).store).textContent).toContain('未提交的正文内容');
 expect(new SurfaceStore((await target.openDocument(restored.id)).store).list()[0].text).toBe('白板文字');
 expect(target.search.search('未提交的正文内容').map(hit=>hit.id)).toEqual([restored.id]);
 await ws.destroy();await target.destroy();
});

it('单文档草稿可解析恢复；附件缺失在预览与恢复结果中提示',async()=>{
 const ws=new LocalWorkspace(new IndexedDBStorage(crypto.randomUUID()),'local',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 await ws.open();
 const id=await ws.create('带附件草稿','正文');
 const item=await ws.openDocument(id);
 const asset=await ws.assets.save(new File(['附件内容'],'文件.txt',{type:'text/plain'}),false);
 writeMarkdown(item.store,`正文与附件\n\n[文件](asset:${asset.id})`);
 await item.persistence.flush();
 // 单文档草稿格式与界面导出一致：format/version/id/update。
 const draft={format:'mydoc-document-draft',version:1,id,update:Array.from(Y.encodeStateAsUpdate(item.store.doc))};

 const preview=parseDraft(JSON.stringify(draft));
 expect(preview.kind).toBe('document');
 expect(preview.documents).toHaveLength(1);
 expect(preview.documents[0].missingAssets).toEqual([asset.id]);

 // 目标工作区没有该附件字节：恢复成功但明确提示缺失。
 const target=new LocalWorkspace(new IndexedDBStorage(crypto.randomUUID()),'local',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 await target.open();
 const [result]=await restoreDraft(target,preview.documents);
 expect(result.missingAssets).toEqual([asset.id]);
 expect(readDocument((await target.openDocument(result.id)).store).textContent).toContain('正文与附件');

 // 目标工作区已有同 ID 附件时，不再报缺失。
 const withAsset=new LocalWorkspace(new IndexedDBStorage(crypto.randomUUID()),'local',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 await withAsset.open();
 await withAsset.assets.put({id:asset.id,name:'文件.txt',mime:'text/plain',size:4,bytes:new Uint8Array([1,2,3,4])});
 const [kept]=await restoreDraft(withAsset,preview.documents);
 expect(kept.missingAssets).toEqual([]);
 await ws.destroy();await target.destroy();await withAsset.destroy();
});

it('拒绝损坏、超限与不支持的草稿；团队工作空间不能作为恢复目标',async()=>{
 const ws=new LocalWorkspace(new IndexedDBStorage(crypto.randomUUID()),'local',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 await ws.open();
 const id=await ws.create('内容','正文');
 const draft=await ws.exportDraft();

 expect(()=>parseDraft('{')).toThrow('JSON');
 expect(()=>parseDraft(JSON.stringify({format:'mydoc-workspace-draft',version:2,workspaceId:'local',documents:[]}))).toThrow('版本');
 expect(()=>parseDraft(JSON.stringify({format:'other-tool-draft',version:1}))).toThrow('不是支持的草稿文件');
 expect(()=>parseDraft(JSON.stringify({...draft,workspaceId:''}))).toThrow('工作空间标识无效');
 expect(()=>parseDraft(JSON.stringify({...draft,createdAt:'bad'}))).toThrow('创建时间无效');
 // 没有已打开文档时导出的草稿只有 ID 字符串，没有可恢复内容。
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[id]}))).toThrow('没有可恢复的文档内容');
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[{id:'',update:[]}]}))).toThrow('草稿文档信息无效');
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[{id:'a',title:5,update:[]}]}))).toThrow('标题无效');
 const entry=draft.documents[0];
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[entry,{...entry}]}))).toThrow('重复的文档');
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[{id:'a',update:'not-array'}]}))).toThrow('数据无效');
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[{id:'a',update:[1,2,999]}]}))).toThrow('数据无效');
 expect(()=>parseDraft(JSON.stringify({...draft,documents:[{id:'a',update:[255,255,255,255]}]}))).toThrow('损坏');
 expect(()=>parseDraft(JSON.stringify({format:'mydoc-document-draft',version:1,id:'',update:[]}))).toThrow('文档标识无效');

 // 团队工作空间是同步目标，不能作为草稿恢复目标：避免绕过 tombstone 或重新发布内容。
 const cloud=new LocalWorkspace(new CloudStorage('ws-1'),'ws-1',true,new IndexedDBAssetStorage(crypto.randomUUID()));
 const preview=parseDraft(JSON.stringify(draft));
 await expect(restoreDraft(cloud,preview.documents)).rejects.toThrow('只能恢复到个人工作区');
 await expect(restoreDraft(ws,[])).rejects.toThrow('请选择要恢复的文档');
 cloud.sync?.close();
 await ws.destroy();
});
