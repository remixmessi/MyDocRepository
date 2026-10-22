import {CLOUD_ORIGIN,registerLocalOffice,CloudStorage,CloudAssets,cloudRequest,cloudJson,cloudServerUrl,type CloudWorkspace,type CloudAccount} from '../../../common/nbstore/src/cloud';
import { templateStore, disposeTemplate, snapshotAssets } from './document-templates';
import { encodeBinary } from './backup';
import { Y, LOCAL_ORIGIN } from '@mydoc/crdt';
import { PersistentDocument, IndexedDBAssetStorage, type AssetStorage, type DocumentStorage } from '@mydoc/nbstore';
import { remapReferences, remapAssets } from '@mydoc/editor-blocks';
import { LocalWorkspace } from './workspace';
import { decodeBinary, type WorkspaceBackup } from './backup';
export interface WorkspaceMeta { id: string; name: string; createdAt: number; cloud?:boolean; userId?:string;role?:CloudWorkspace['role'] }
export class WorkspaceRegistry {
  private pendingPersonalRestore?:{source:string;id:string;operation:string};
  private server=cloudServerUrl()||globalThis.location?.origin||'http://localhost';
  private doc = new Y.Doc(); readonly entries = this.doc.getMap<WorkspaceMeta>('workspaces'); private preferences = this.doc.getMap<string>('preferences');
  readonly persistence: PersistentDocument; private opened = new Map<string, Promise<LocalWorkspace>>();
  constructor(private storage: DocumentStorage, private assets: AssetStorage = new IndexedDBAssetStorage()) { this.persistence = new PersistentDocument('registry:local', this.doc, storage); }
  async open() { await this.persistence.open();
    const previous=this.preferences.get('cloud-registry-server');
    if(previous&&previous!==this.server){
      const clouds=this.list().filter(entry=>entry.cloud);if(clouds.length&&previous)this.preferences.set('cloud-spaces:'+previous,JSON.stringify(clouds));
      let restored:WorkspaceMeta[]=[];try{restored=JSON.parse(this.preferences.get('cloud-spaces:'+this.server)??'[]');}catch{}
      this.doc.transact(()=>{for(const entry of clouds)this.entries.delete(entry.id);for(const entry of restored)if(entry.cloud)this.entries.set(entry.id,entry);this.preferences.set('cloud-registry-server',this.server);const local=this.list().find(entry=>!entry.cloud);if(local)this.preferences.set('active',local.id);},LOCAL_ORIGIN);
      await this.persistence.flush();
    }
    if(!previous){this.preferences.set('cloud-registry-server',this.server);this.preferences.set('cloud-spaces:'+this.server,JSON.stringify(this.list().filter(entry=>entry.cloud)));await this.persistence.flush();}
    if (!this.entries.size) { this.doc.transact(() => { this.entries.set('local', { id: 'local', name: '本地工作区', createdAt: Date.now() }); this.preferences.set('active', 'local'); }, LOCAL_ORIGIN); await this.persistence.flush(); } }
  subscribe = (listener: () => void) => { this.entries.observe(listener); this.preferences.observe(listener); return () => { this.entries.unobserve(listener); this.preferences.unobserve(listener); }; };
  list() { return [...this.entries.values()].sort((a, b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id)); }
  get activeId() { const saved = this.preferences.get('active'); return saved && this.entries.has(saved) ? saved : this.list()[0]?.id ?? 'local'; }
  async get(id: string) { if (!this.entries.has(id)) throw new Error('工作区不存在'); let promise = this.opened.get(id); if (!promise) { const ws = new LocalWorkspace(this.entries.get(id)!.cloud?new CloudStorage(id,this.entries.get(id)!.role):this.storage, id, !!this.entries.get(id)!.cloud, this.entries.get(id)!.cloud?new CloudAssets(id):this.assets,this.entries.get(id)!.userId,this.storage); promise = ws.open().then(() => ws).catch(async error => { await ws.destroy(); this.opened.delete(id); throw error; }); this.opened.set(id, promise); } return promise; }
  async select(id: string) { const next = await this.get(id); this.doc.transact(() => this.preferences.set('active', id), LOCAL_ORIGIN); await this.persistence.flush(); return next; }
  async create(name: string) { if (!name.trim()) throw new Error('请输入工作区名称'); const id = crypto.randomUUID(); this.doc.transact(() => this.entries.set(id, { id, name: name.trim(), createdAt: Date.now() }), LOCAL_ORIGIN); try { await this.persistence.flush(); } catch (error) { this.doc.transact(() => this.entries.delete(id), LOCAL_ORIGIN); throw error; } return this.select(id); }
  /**
   * Rebind the registry to the currently configured cloud server without
   * reloading the application. Cloud workspaces are cached per server so a
   * server switch never replaces one server's directory with another's.
   */
  async switchServer(server=cloudServerUrl()||globalThis.location?.origin||'http://localhost') {
    const next=server.trim().replace(/\/+$/,'')||'http://localhost';
    if(next===this.server)return;
    const previous=this.server;
    const clouds=this.list().filter(entry=>entry.cloud);
    if(clouds.length)this.preferences.set('cloud-spaces:'+previous,JSON.stringify(clouds));
    for(const [id,promise] of this.opened){
      if(this.entries.get(id)?.cloud){try{await (await promise).destroy();}finally{this.opened.delete(id);}}
    }
    let restored:WorkspaceMeta[]=[];
    try{restored=JSON.parse(this.preferences.get('cloud-spaces:'+next)??'[]');}catch{}
    this.server=next;
    this.doc.transact(()=>{
      for(const entry of clouds)this.entries.delete(entry.id);
      for(const entry of restored)if(entry.cloud)this.entries.set(entry.id,entry);
      this.preferences.set('cloud-registry-server',next);
      const active=this.preferences.get('active');
      if(!active||!this.entries.has(active)||this.entries.get(active)?.cloud)this.preferences.set('active',this.list().find(entry=>!entry.cloud)?.id??'local');
    },LOCAL_ORIGIN);
    await this.persistence.flush();
  }
  async refreshCloud(){
    const {user}=await cloudJson<{user:CloudAccount|null}>('/session');
    if(!user){
      const active=this.preferences.get('active');
      const clouds=this.list().filter(entry=>entry.cloud);
      for(const [id,promise] of this.opened){
        if(this.entries.get(id)?.cloud){await (await promise).destroy();this.opened.delete(id);}
      }
      this.doc.transact(()=>{
        for(const entry of clouds)this.entries.delete(entry.id);
        if(active&&clouds.some(entry=>entry.id===active))this.preferences.set('active',this.list().find(entry=>!entry.cloud)?.id??'local');
        this.preferences.set('cloud-spaces:'+this.server,'[]');
      },LOCAL_ORIGIN);
      await this.persistence.flush();
      return;
    }
    const spaces=await cloudJson<CloudWorkspace[]>('/workspaces');
    for(const [id,promise]of this.opened)if(this.entries.get(id)?.cloud&&!spaces.some(space=>space.id===id)){const ws=await promise;if(ws.sync?.status==='blocked'||ws.persistence.state==='error')throw new Error('工作空间权限已失效，请先在当前页面导出草稿。');}
    const active=this.preferences.get('active');
    this.doc.transact(()=>{
      for(const entry of this.list())if(entry.cloud&&!spaces.some(s=>s.id===entry.id))this.entries.delete(entry.id);
      for(const space of spaces)this.entries.set(space.id,{id:space.id,name:space.name,createdAt:space.createdAt,cloud:true,userId:user.id,role:space.role});
      if(active&&this.entries.get(active)===undefined)this.preferences.set('active',this.list().find(entry=>!entry.cloud)?.id??'local');
      this.preferences.set('cloud-spaces:'+this.server,JSON.stringify(this.list().filter(entry=>entry.cloud)));
    },LOCAL_ORIGIN);
    await this.persistence.flush();
    for(const[id,promise]of this.opened)if(!this.entries.has(id)){await (await promise).destroy();this.opened.delete(id);}
  }
  async rename(id: string, name: string) { const entry = this.entries.get(id); if(entry?.cloud)throw new Error('请在团队工作空间设置中修改名称'); if (!entry || !name.trim()) throw new Error('请输入工作区名称'); this.doc.transact(() => this.entries.set(id, { ...entry, name: name.trim() }), LOCAL_ORIGIN); await this.persistence.flush(); }
  async delete(id: string) {
    if(this.entries.get(id)?.cloud)throw new Error('请在团队工作空间设置中删除工作空间');
    if (this.entries.size < 2) throw new Error('至少保留一个工作区'); const entry = this.entries.get(id)!; const previousActive = this.activeId; const ws = await this.get(id); await ws.flush();
    this.doc.transact(() => { this.entries.delete(id); if (this.activeId === id || this.preferences.get('active') === id) this.preferences.set('active', this.list()[0].id); }, LOCAL_ORIGIN); try { await this.persistence.flush(); } catch (error) { this.doc.transact(() => { this.entries.set(id, entry); this.preferences.set('active', previousActive); }, LOCAL_ORIGIN); throw error; }
    await ws.purgeAllDocuments(); await ws.assets.clear(); await ws.destroy(); this.opened.delete(id); await this.storage.remove(`workspace:${id}`);
  }
  async restore(backup: WorkspaceBackup) {
    if(backup.office){
      const source=JSON.stringify(backup);if(!this.pendingPersonalRestore||this.pendingPersonalRestore.source!==source)this.pendingPersonalRestore={source,id:crypto.randomUUID(),operation:crypto.randomUUID()};const pending=this.pendingPersonalRestore;await this.flush();const unregister=registerLocalOffice(pending.id);
      try{const response=await cloudRequest('/office-workspace-full-restore?'+new URLSearchParams({workspace:pending.id,operationId:pending.operation}),{method:'POST',headers:{'Content-Type':'application/json'},body:source}),result=await response.json() as {id:string;registryUpdate:string};Y.applyUpdate(this.doc,decodeBinary(result.registryUpdate),CLOUD_ORIGIN);const ws=await this.select(result.id);this.pendingPersonalRestore=undefined;return ws;}finally{unregister();}
    }
    const previousActive = this.activeId; const id = crypto.randomUUID(); const ws = new LocalWorkspace(this.storage, id, false, this.assets); await ws.open(); const folderIds = new Map(backup.folders.map(folder => [folder.id, crypto.randomUUID()]));
    try {
      for (const asset of backup.assets ?? []) await ws.assets.put({ ...asset, bytes: asset.data ? decodeBinary(asset.data) : new Uint8Array() });
      ws.folders.doc!.transact(() => backup.folders.forEach(folder => ws.folders.set(folderIds.get(folder.id)!, { ...folder, id: folderIds.get(folder.id)!, parentId: folder.parentId ? folderIds.get(folder.parentId)! : null })), LOCAL_ORIGIN);
      const collectionIds = new Map((backup.collections??[]).map(c=>[c.id,crypto.randomUUID()]));
      const entries=[...backup.documents,...(backup.trash??[])];const backupAssets=new Map((backup.assets??[]).map(asset=>[asset.id,asset]));
      const docIds = new Map<string,string>();
      const restoreEntry=async(entry:typeof backup.documents[number],recycled?:{deletedAt:number;collectionIds:string[]})=>{const target=crypto.randomUUID(),store=templateStore({id:entry.meta.id,name:entry.meta.title,createdAt:entry.meta.createdAt,sourceDocId:entry.meta.id,sourceWorkspaceId:backup.workspaceId??'local',update:entry.update});try{for(const assetId of snapshotAssets(store)){const asset=backupAssets.get(assetId);if(!asset)throw new Error('备份附件缺失');await ws.documentAssets(target).put({...asset,id:assetId,bytes:asset.data?decodeBinary(asset.data):new Uint8Array()});}const result=await ws.createSnapshot(entry.meta.title,entry.meta.createdAt,entry.meta.folderId?folderIds.get(entry.meta.folderId)!:null,Y.encodeStateAsUpdate(store.doc),recycled,undefined,target);return result;}finally{disposeTemplate(store);}};
      for(const entry of backup.documents)docIds.set(entry.meta.id,await restoreEntry(entry));
      for(const entry of backup.trash??[])docIds.set(entry.meta.id,await restoreEntry(entry,{deletedAt:entry.deletedAt,collectionIds:entry.collectionIds.map(key=>collectionIds.get(key)!)}));
      const tagIds = new Map((backup.tags ?? []).map(tag => [tag.id, crypto.randomUUID()]));
      ws.documents.doc!.transact(() => {
        for (const tag of backup.tags ?? []) ws.organization.tags.set(tagIds.get(tag.id)!, { ...tag, id: tagIds.get(tag.id)! });
        for (const entry of entries) { const newId = docIds.get(entry.meta.id)!; const meta={...(ws.documents.get(newId)??ws.trash.get(newId)!.meta),tagIds:(entry.meta.tagIds??[]).map(id=>tagIds.get(id)!),favorite:entry.meta.favorite??false,...(entry.meta.journalDate?{journalDate:entry.meta.journalDate}:{}),...(entry.meta.webClip?{webClip:entry.meta.webClip}:{}),...(entry.meta.importSource?{importSource:entry.meta.importSource}:{})};const trash=ws.trash.get(newId);if(trash)ws.trash.set(newId,{...trash,meta});else ws.documents.set(newId,meta); }
        for(const entry of backup.officeOrganization??[]){const {id:resourceId,...value}=entry;ws.officeOrganization.set(resourceId,{...value,folderId:value.folderId?folderIds.get(value.folderId)!:null,tagIds:(value.tagIds??[]).map(id=>tagIds.get(id)!)});}
        for (const collection of backup.collections ?? []) { const key = collectionIds.get(collection.id)!; ws.organization.collections.set(key, { ...collection, id: key, documentIds: collection.documentIds.map(id => docIds.get(id)!), rules: { ...collection.rules, tagIds: collection.rules.tagIds.map(id => tagIds.get(id)!) } }); }
      }, LOCAL_ORIGIN);
      for(const template of backup.templates??[]){const store=templateStore(template);try{remapReferences(store,backup.workspaceId||'local',id,docIds);const next={...template,id:crypto.randomUUID(),sourceWorkspaceId:template.sourceWorkspaceId===(backup.workspaceId||'local')?id:template.sourceWorkspaceId,sourceDocId:template.sourceWorkspaceId===(backup.workspaceId||'local')?(docIds.get(template.sourceDocId)??template.sourceDocId):template.sourceDocId,update:encodeBinary(Y.encodeStateAsUpdate(store.doc))};ws.templates.set(next.id,next);}finally{disposeTemplate(store);}}
      for (const newId of docIds.values()) if(!ws.isOfficeResource(newId))remapReferences((await ws.openDocument(newId)).store, backup.workspaceId || 'local', id, docIds); await ws.flush();
      this.doc.transact(() => { this.entries.set(id, { id, name: `${backup.name}（恢复）`, createdAt: Date.now() }); this.preferences.set('active', id); }, LOCAL_ORIGIN); await this.persistence.flush(); this.opened.set(id, Promise.resolve(ws)); return ws;
    } catch (error) { this.doc.transact(() => { this.entries.delete(id); this.preferences.set('active', previousActive); }, LOCAL_ORIGIN); await this.persistence.flush().catch(()=>{}); await ws.purgeAllDocuments().catch(() => {}); await ws.assets.clear().catch(() => {}); await ws.destroy(); await this.storage.remove(`workspace:${id}`).catch(() => {}); throw error; }
  }
  async flush() { for(const promise of this.opened.values()){const ws=await promise;if(ws.assets.saving)throw new Error('附件正在上传，请完成后再关闭。');await ws.flush();}await this.persistence.flush(); }
  async destroy() { for (const promise of this.opened.values()) { try { await (await promise).destroy(); } catch { /* failed opens are disposed */ } } await this.persistence.destroy(); this.storage.close(); this.assets.close(); }
}
