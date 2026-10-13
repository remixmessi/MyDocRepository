import {CloudStorage} from '../../../common/nbstore/src/cloud';
import {validWebClip} from './clip-source';
import { validJournalDate, journalTitle, journalMarkdown, type JournalTemplate } from './journal';
import { Y, LOCAL_ORIGIN } from '@mydoc/crdt';
import { BlockStore } from '@mydoc/editor-framework';
import { readDocument, markdownSerializer, writeMarkdown, writeDocument, documentText, noteText, surfaceReferences, validSurfaceElement, type EmbedSource } from '@mydoc/editor-blocks';
import { PersistentDocument, WorkspaceAssets, IndexedDBAssetStorage, type AssetStorage, type DocumentStorage } from '@mydoc/nbstore';
import type { AssetRecord } from '@mydoc/nbstore';
import type { DocumentTemplate } from './document-templates';
import { DocumentSearch } from './search';
import { ReferenceIndex } from './references';
import { DocumentOrganization } from './organization';

export interface DocumentMeta { id: string; title: string; createdAt: number; folderId?: string | null; tagIds?: string[]; favorite?: boolean; journalDate?: string; importSource?: import('./document-import').ImportSource; webClip?: import('./clip-source').WebClipSource }
export interface TrashEntry { meta: DocumentMeta; deletedAt: number; collectionIds: string[] }
export interface FolderMeta { id: string; name: string; parentId: string | null; createdAt: number }
export interface OpenDocument { store: BlockStore; persistence: PersistentDocument }
class DocumentAssetStorage implements AssetStorage {
  constructor(private base:AssetStorage,private doc:string) {}
  get(workspace:string,id:string){return this.base.getForDocument?.(workspace,id,this.doc)??this.base.get(workspace,id);}
  put(workspace:string,asset:AssetRecord){return this.base.putForDocument?.(workspace,asset,this.doc)??this.base.put(workspace,asset);}
  remove(workspace:string,id:string){return this.base.removeForDocument?.(workspace,id,this.doc)??this.base.remove?.(workspace,id)??Promise.resolve();}
  removeWorkspace(workspace:string){return this.base.removeWorkspace(workspace);}
  close(){this.base.close();}
}
export class LocalWorkspace {
  private journalRequests = new Map<string, Promise<string>>();
  private doc = new Y.Doc();
  readonly documents = this.doc.getMap<DocumentMeta>('documents');
  readonly templates = this.doc.getMap<DocumentTemplate>('documentTemplates');
  readonly trash = this.doc.getMap<TrashEntry>('trash');
  readonly uiPreferences:Y.Map<unknown>;private personalPersistence?:PersistentDocument;
  readonly folders = this.doc.getMap<FolderMeta>('folders');
  readonly organization = new DocumentOrganization(this.doc, this.documents, id => this.search.text(id),()=>this.assertWritable());
  private hiddenDocuments = new Set<string>();
  readonly persistence: PersistentDocument; readonly assets: WorkspaceAssets; private assetStorage: AssetStorage; private ownsAssets: boolean;
  readonly search = new DocumentSearch(); readonly references = new ReferenceIndex();
  indexReady: Promise<void> = Promise.resolve(); indexing = false; indexErrors = new Set<string>();
  private openedItems = new Map<string,OpenDocument>();
  private documentAssetViews = new Map<string,WorkspaceAssets>();
  private opened = new Map<string, Promise<OpenDocument>>(); private listeners = new Set<() => void>(); private indexTimers = new Map<string, ReturnType<typeof setTimeout>>(); private disposed = false;
  constructor(private storage: DocumentStorage, readonly id = 'local', private ownsStorage = true, assets?: AssetStorage,userId?:string,personalStorage?:DocumentStorage) {
    if(userId&&personalStorage){const personalDoc=new Y.Doc();this.uiPreferences=personalDoc.getMap<unknown>('uiPreferences');this.personalPersistence=new PersistentDocument('workspace-ui:'+this.id,personalDoc,personalStorage);}else this.uiPreferences=this.doc.getMap<unknown>('uiPreferences');
    this.ownsAssets = !assets; this.assetStorage = assets ?? new IndexedDBAssetStorage(); this.assets = new WorkspaceAssets(this.assetStorage, id);
    this.persistence = new PersistentDocument(`workspace:${id}`, this.doc, storage);
    this.uiPreferences.observe(event=>{if(event.keysChanged.has('editorSettings'))this.notify();}); this.templates.observe(this.notify); this.documents.observe(this.notify); this.trash.observe(this.notify); this.folders.observe(this.notify); this.organization.tags.observe(()=>{if(this.readOnly){this.notify();return;}const tags=new Set(this.organization.tags.keys());this.doc.transact(()=>{for(const [id,entry]of this.trash){const tagIds=(entry.meta.tagIds??[]).filter(id=>tags.has(id));if(tagIds.length!==(entry.meta.tagIds??[]).length)this.trash.set(id,{...entry,meta:{...entry.meta,tagIds}});}},LOCAL_ORIGIN);this.notify();}); this.organization.collections.observe(()=>{if(this.readOnly){this.notify();return;}const keys=new Set(this.organization.collections.keys());this.doc.transact(()=>{for(const[id,entry]of this.trash){const collectionIds=entry.collectionIds.filter(id=>keys.has(id));if(collectionIds.length!==entry.collectionIds.length)this.trash.set(id,{...entry,collectionIds});}},LOCAL_ORIGIN);this.notify();});
  }
  get sync(){return this.storage instanceof CloudStorage?this.storage:undefined;}
  documentAssets(id:string){
    let assets=this.documentAssetViews.get(id);
    if(!assets){
      assets=new WorkspaceAssets(new DocumentAssetStorage(this.assetStorage,'doc:'+id),this.id);
      this.documentAssetViews.set(id,assets);
    }
    return assets;
  }
  get readOnly(){return this.sync?.readOnly??false;}
  documentReadOnly(id:string){return this.sync?.isReadOnly('doc:'+id)??false;}
  assertWritable(id?:string){if(this.readOnly||id&&this.documentReadOnly(id))throw new Error(this.sync?.reason||'该文档为只读');}
  draft(){return {format:'mydoc-workspace-draft',version:1,workspaceId:this.id,createdAt:Date.now(),workspace:Array.from(Y.encodeStateAsUpdate(this.doc)),documents:[...this.opened].map(([id])=>id)};}
  get pendingDraft(){return [...this.openedItems.values()].some(item=>item.persistence.state==='saving'||item.persistence.state==='error')||this.persistence.state==='saving'||this.persistence.state==='error'||this.sync?.status==='offline'||this.sync?.status==='blocked';}
  async exportDraft(){const documents=[];for(const [id,promise]of this.opened){let item:OpenDocument;try{item=await promise;}catch{continue;}documents.push({id,title:item.store.meta.get('title'),markdown:markdownSerializer.serialize(readDocument(item.store)),update:Array.from(Y.encodeStateAsUpdate(item.store.doc))});}return {...this.draft(),documents};}
  readonly embedSource: EmbedSource = { load: async (id, workspaceId) => { if (workspaceId !== this.id || !this.isDocumentVisible(id)) return null; const item = await this.openDocument(id); if (!this.isDocumentVisible(id)) return null; return { title: this.documents.get(id)!.title, document: readDocument(item.store), surface:[...item.store.doc.getMap('surface').values()].filter(validSurfaceElement), assets: this.documentAssets(id) }; }, subscribe: callback => this.subscribe(callback) };
  private notify = () => this.listeners.forEach(listener => listener());
  async open() {
    await this.persistence.open();if(this.sync){try{this.hiddenDocuments=new Set(this.documents.keys());const visible=await this.sync.visibleDocuments([...this.documents.keys()]);this.hiddenDocuments=new Set([...this.documents.keys()].filter(id=>!visible.has('doc:'+id)));}catch{this.hiddenDocuments.clear();}}this.sync?.subscribe(this.notify);await this.personalPersistence?.open(); this.indexing = true;
    this.indexReady = Promise.allSettled(this.list().map(async meta => this.indexDocument(meta.id, await this.openDocument(meta.id).catch(error => { this.indexErrors.add(meta.id); throw error; })))).then(() => { this.indexing = false; this.notify(); });
  }
  subscribe = (callback: () => void) => { this.listeners.add(callback); return () => { this.listeners.delete(callback); }; };
  isDocumentVisible(id:string){return !this.hiddenDocuments.has(id);}
  list() { return [...this.documents.values()].filter(meta=>this.isDocumentVisible(meta.id)).sort((a, b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id)); }
  trashList() { return [...this.trash.values()].sort((a,b)=>b.deletedAt-a.deletedAt || a.meta.id.localeCompare(b.meta.id)); }
  allDocuments() { return [...this.list(), ...this.trashList().map(entry=>entry.meta)]; }
  folderList() { return [...this.folders.values()].sort((a, b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id)); }
  private indexDocument(id: string, item: OpenDocument) {
    if (!this.documents.has(id) || this.disposed) return;
    try { const node = readDocument(item.store); this.search.update({ id, title: this.documents.get(id)!.title, text: (this.documents.get(id)!.journalDate??'') + '\n' + documentText(node, (target, ws) => ws === this.id ? this.documents.get(target)?.title : undefined) + '\n' + [...item.store.doc.getMap('surface').values()].filter(validSurfaceElement).map(shape => shape.kind==='note'?noteText(shape.markdown,(target,ws)=>ws===this.id?this.documents.get(target)?.title:undefined):shape.kind==='doc-card'?(shape.workspaceId===this.id?this.documents.get(shape.docId)?.title||shape.text:shape.text):shape.kind==='mindmap'?shape.nodes.map(n=>n.text).join('\n'):shape.text).join('\n') }); this.references.update(id, node, (target, ws) => ws === this.id ? this.documents.get(target)?.title : undefined, surfaceReferences([...item.store.doc.getMap('surface').values()].filter(validSurfaceElement), (target,ws)=>ws===this.id?this.documents.get(target)?.title:undefined)); this.indexErrors.delete(id); }
    catch { this.indexErrors.add(id); }
    this.notify();
  }
  private scheduleIndex(id: string, item: OpenDocument) {
    if (this.disposed) return; clearTimeout(this.indexTimers.get(id)); this.indexTimers.set(id, setTimeout(() => { this.indexTimers.delete(id); this.indexDocument(id, item); }, 80));
  }
  journal(date:string) { return this.list().find(meta=>meta.journalDate===date); }
  recycledJournal(date:string) { return this.trashList().find(entry=>entry.meta.journalDate===date); }
  getOrCreateJournal(date:string, template:JournalTemplate='daily'):Promise<string> {
    if(!validJournalDate(date))return Promise.reject(new Error('请选择有效日期'));
    const pending=this.journalRequests.get(date);if(pending)return pending;
    const existing=this.journal(date);if(existing)return Promise.resolve(existing.id);
    if(this.recycledJournal(date))return Promise.reject(new Error('该日期的日记在回收站中，请恢复原日记'));
    const job=(async()=>{if(this.persistence.state==='error')await this.persistence.retry();return this.create(journalTitle(date),journalMarkdown(date,template),null,undefined,date);})();
    this.journalRequests.set(date,job);void job.finally(()=>this.journalRequests.delete(date)).catch(()=>{});return job;
  }
  async create(title = '无标题文档', markdown?: string, folderId: string | null = null, imported?:{content:ReturnType<typeof readDocument>;properties?:Pick<DocumentMeta,'webClip'|'tagIds'>}, journalDate?:string) { this.assertWritable();
    if(journalDate!==undefined&&(!validJournalDate(journalDate)||this.allDocuments().some(meta=>meta.journalDate===journalDate)))throw new Error('日记日期无效或已经存在');
    if(imported?.properties?.webClip&&!validWebClip(imported.properties.webClip))throw new Error('剪藏来源无效');
    if(imported?.properties?.tagIds?.some(id=>!this.organization.tags.has(id)))throw new Error('标签不存在');
    if (folderId && !this.folders.has(folderId)) throw new Error('文件夹不存在');
    const id = crypto.randomUUID(); const item = await this.openDocument(id, { create: true });
    try { if(imported)writeDocument(item.store,imported.content);else if (markdown !== undefined) writeMarkdown(item.store, markdown); item.store.rename(title);
    await item.persistence.flush();
    this.doc.transact(() => this.documents.set(id, { id, title, createdAt: Date.now(), folderId, ...imported?.properties, ...(journalDate?{journalDate}:{}) }), LOCAL_ORIGIN);
    await this.persistence.flush(); item.store.history.clear(); this.indexDocument(id, item); return id; } catch (error) { await this.discardUnpublished(id); throw error; }
  }
  async createSnapshot(title: string, createdAt: number, folderId: string | null, update: Uint8Array, recycled?: { deletedAt: number; collectionIds: string[] }, transform?: (store: BlockStore, id: string) => void, forcedId?: string) { this.assertWritable();
    const id = forcedId ?? crypto.randomUUID(); try { await this.storage.append(`doc:${id}`, update); const item = await this.openDocument(id);
    readDocument(item.store); transform?.(item.store,id); item.store.rename(title); await item.persistence.flush();
    this.doc.transact(() => { const meta={ id, title, createdAt, folderId }; if(recycled)this.trash.set(id,{meta,...recycled});else this.documents.set(id,meta); }, LOCAL_ORIGIN); await this.persistence.flush(); item.store.history.clear(); this.indexDocument(id, item); return id; } catch (error) { await this.discardUnpublished(id); throw error; }
  }
  private async discardUnpublished(id: string) {
    this.doc.transact(() => {this.documents.delete(id);this.trash.delete(id);}, LOCAL_ORIGIN); this.search.remove(id); this.references.remove(id); clearTimeout(this.indexTimers.get(id)); this.indexTimers.delete(id);
    const opened = this.opened.get(id); if (opened) { try { const item = await opened; item.store.destroy(); await item.persistence.destroy(); } catch { /* failed opens dispose themselves */ } this.opened.delete(id);this.openedItems.delete(id); }
    await this.storage.remove(`doc:${id}`).catch(() => {});
  }
  openDocument(id: string, options: { create?: boolean } = {}): Promise<OpenDocument> {
    if(!options.create&&!this.isDocumentVisible(id))return Promise.reject(new Error('文档不存在或无权访问'));
    let promise = this.opened.get(id);
    if (!promise) {
      promise = (async () => {
        const doc = new Y.Doc(); const persistence = new PersistentDocument(`doc:${id}`, doc, this.storage);
        try {
          // A freshly generated ID has no stored content to read yet.
          await persistence.open({ load: !options.create }); const store = new BlockStore(doc); store.initialize(); store.ensureNote();
          const item = { store, persistence };this.openedItems.set(id,item);
          const syncTitle = () => { const entry = this.documents.get(id); const title = store.meta.get('title') ?? '无标题文档'; if (!this.documentReadOnly(id) && entry && entry.title !== title) { this.doc.transact(() => this.documents.set(id, { ...entry, title }), LOCAL_ORIGIN); for (const link of this.references.backlinks(id, this.id)) void this.opened.get(link.sourceId)?.then(source => this.scheduleIndex(link.sourceId, source)); } };
          store.meta.observe(syncTitle); syncTitle(); store.doc.on('update', () => this.scheduleIndex(id, item)); return item;
        } catch (error) { await persistence.destroy(); throw error; }
      })();
      this.opened.set(id, promise); void promise.catch(() => this.opened.delete(id));
    }
    return promise;
  }
  rename(id: string, title: string, store: BlockStore) { this.assertWritable(id); if (this.documents.has(id)) store.rename(title); }
  createFolder(name: string, parentId: string | null = null) { this.assertWritable();
    if (!name.trim()) throw new Error('请输入文件夹名称'); if (parentId && !this.folders.has(parentId)) throw new Error('上级文件夹不存在');
    const id = crypto.randomUUID(); this.doc.transact(() => this.folders.set(id, { id, name: name.trim(), parentId, createdAt: Date.now() }), LOCAL_ORIGIN); return id;
  }
  renameFolder(id: string, name: string) { this.assertWritable(); const folder = this.folders.get(id); if (!folder || !name.trim()) throw new Error('请输入文件夹名称'); this.doc.transact(() => this.folders.set(id, { ...folder, name: name.trim() }), LOCAL_ORIGIN); }
  moveFolder(id: string, parentId: string | null) { this.assertWritable();
    const folder = this.folders.get(id); if (!folder) throw new Error('文件夹不存在'); let parent = parentId;
    const seen = new Set<string>([id]); while (parent) { if (seen.has(parent)) throw new Error('不能将文件夹移动到自身或其子文件夹'); seen.add(parent); const next = this.folders.get(parent); if (!next) throw new Error('上级文件夹不存在'); parent = next.parentId; }
    this.doc.transact(() => this.folders.set(id, { ...folder, parentId }), LOCAL_ORIGIN);
  }
  moveDocument(id: string, folderId: string | null) { this.assertWritable(id); const meta = this.documents.get(id); if (!meta) throw new Error('文档不存在'); if (folderId && !this.folders.has(folderId)) throw new Error('文件夹不存在'); this.doc.transact(() => this.documents.set(id, { ...meta, folderId }), LOCAL_ORIGIN); }
  deleteFolder(id: string) { this.assertWritable();
    const folder = this.folders.get(id); if (!folder) return;
    this.doc.transact(() => { for (const doc of this.list()) if (doc.folderId === id) this.documents.set(doc.id, { ...doc, folderId: folder.parentId }); for (const entry of this.trashList()) if(entry.meta.folderId===id)this.trash.set(entry.meta.id,{...entry,meta:{...entry.meta,folderId:folder.parentId}}); for (const child of this.folderList()) if (child.parentId === id) this.folders.set(child.id, { ...child, parentId: folder.parentId }); this.folders.delete(id); }, LOCAL_ORIGIN);
  }
  private refreshReferences() { for (const [id,promise] of this.opened) if(this.documents.has(id)) void promise.then(item=>this.scheduleIndex(id,item)).catch(()=>{}); }
  private unindex(id: string) { clearTimeout(this.indexTimers.get(id)); this.indexTimers.delete(id); this.search.remove(id); this.references.remove(id); this.indexErrors.delete(id); }
  async delete(id: string) { this.assertWritable(id);
    const meta=this.documents.get(id); if(!meta)return; const item=await this.openDocument(id); await item.persistence.flush();
    const collectionIds=this.organization.collectionList().filter(c=>c.mode==='manual'&&c.documentIds.includes(id)).map(c=>c.id);
    this.doc.transact(()=>{this.trash.set(id,{meta,deletedAt:Date.now(),collectionIds});this.organization.removeDocument(id);this.documents.delete(id);},LOCAL_ORIGIN);this.unindex(id);
    try { await this.persistence.flush(); } catch(error) { this.doc.transact(()=>{this.trash.delete(id);this.documents.set(id,meta);for(const key of collectionIds)if(this.organization.collections.get(key)?.mode==='manual')this.organization.setMembership(key,id,true);},LOCAL_ORIGIN);this.indexDocument(id,item);this.refreshReferences();throw error; }
    this.refreshReferences();this.notify();
  }
  async restoreDocument(id: string) { this.assertWritable(id);
    const entry=this.trash.get(id);if(!entry)throw new Error('回收文档不存在');const item=await this.openDocument(id);readDocument(item.store);
    const meta={...entry.meta,folderId:entry.meta.folderId&&this.folders.has(entry.meta.folderId)?entry.meta.folderId:null,tagIds:(entry.meta.tagIds??[]).filter(id=>this.organization.tags.has(id))};
    this.doc.transact(()=>{this.trash.delete(id);this.documents.set(id,meta);for(const key of entry.collectionIds)if(this.organization.collections.get(key)?.mode==='manual')this.organization.setMembership(key,id,true);},LOCAL_ORIGIN);
    try { await this.persistence.flush(); } catch(error) { this.doc.transact(()=>{this.organization.removeDocument(id);this.documents.delete(id);this.trash.set(id,entry);},LOCAL_ORIGIN);this.unindex(id);throw error; }
    this.indexDocument(id,item);this.refreshReferences();return id;
  }
  async permanentlyDelete(id: string) { this.assertWritable(id);
    const entry=this.trash.get(id);if(!entry)throw new Error('仅能彻底删除回收站文档');
    this.doc.transact(()=>this.trash.delete(id),LOCAL_ORIGIN);
    try { await this.persistence.flush(); const opened=this.opened.get(id); if(opened){const item=await opened;await item.persistence.flush();item.store.destroy();await item.persistence.destroy();this.opened.delete(id);this.openedItems.delete(id);}await this.storage.remove(`doc:${id}`); }
    catch(error){this.doc.transact(()=>this.trash.set(id,entry),LOCAL_ORIGIN);throw error;}
    this.unindex(id);this.notify();
  }
  async purgeAllDocuments() { for(const meta of this.list())await this.delete(meta.id);for(const entry of this.trashList())await this.permanentlyDelete(entry.meta.id); }
  async flush() { await this.personalPersistence?.flush(); for (const opened of this.opened.values()) await (await opened).persistence.flush(); await this.persistence.flush(); for (const [id, timer] of [...this.indexTimers]) { clearTimeout(timer); this.indexTimers.delete(id); this.indexDocument(id, await this.openDocument(id)); } }
  async destroy() { await this.personalPersistence?.destroy();
    this.disposed = true; await this.indexReady; this.indexTimers.forEach(timer => clearTimeout(timer)); this.indexTimers.clear();
    for (const opened of this.opened.values()) { try { const item = await opened; item.store.destroy(); await item.persistence.destroy(); } catch { /* failed opens already dispose themselves */ } }
    this.listeners.clear(); this.documentAssetViews.clear(); await this.persistence.destroy(); if (this.ownsStorage) this.storage.close(); if (this.ownsAssets) this.assetStorage.close();
  }
}
