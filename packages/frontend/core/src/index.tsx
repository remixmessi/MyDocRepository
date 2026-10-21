import {downloadTeamWorkspaceBackup,TeamWorkspaceRestoreDialog} from './team-workspace-backup-ui';
import {isCloudWorkspaceArchive,type CloudWorkspaceArchive} from '../../../common/nbstore/src/workspace-archive';
import {NewFileDialog} from './new-file-dialog';
import {Y} from '@mydoc/crdt';
import {cloudServerUrl,type CloudStorage} from '../../../common/nbstore/src/cloud';
export {CloudLogin} from './cloud-ui';
import {CloudDialog} from './cloud-ui';
import {SettingsDialog, type SettingsTab} from './settings-center';
import {useDesktopUpdates,DesktopUpdateDialog} from './desktop-update-ui';
import {DatabaseBackupDialog} from './database-backup-ui';
import {DraftRestoreDialog} from './draft-ui';
import {useDocumentLinkIntake} from './document-link-intake';
import {documentLink} from '../../../common/nbstore/src/deep-link';
import {useFileIntake} from './file-intake';
import {projectPreferences} from '@mydoc/nbstore';
import {DocumentPeekScope} from './peek-view-ui';
import {WebClipperDialog} from './web-clipper-ui';
import {JournalDialog} from './journal-ui';
import {localJournalDate} from './journal';
import {DocumentPermissionDialog} from './document-permission-ui';
export {initializeTheme} from './theme';
import {useAppearance,ThemeDialog} from './theme-ui';
import {EditorSettingsDialog} from './editor-settings-ui';
import { DocumentImportDialog } from './document-import-ui';
import type { ImportKind } from './import-source';
import { rememberPalette, paletteHistory, type PaletteCommand } from './command-palette';
import { DocumentTemplateDialog } from './document-template-ui';
import { useEffect, useReducer, useRef, useState, useId, useLayoutEffect, type CSSProperties } from 'react';
import { FileText, Trash2, Undo2, Redo2, Shield } from 'lucide-react';
import { SQLiteStorage, SQLiteAssetStorage, type PersistentDocument } from '@mydoc/nbstore';
import { DocumentEditor as MarkdownDocumentEditor, readEditorSettings, readDocument, writeDocument, documentOutline, documentSchema, Whiteboard } from '@mydoc/editor-blocks';
import { LocalWorkspace, type DocumentMeta, type OpenDocument, type FolderMeta } from './workspace';
import { WorkspaceRegistry } from './workspaces';
import { DocumentSidebar, FolderSelect } from './sidebar';
import { DocumentOverview } from './document-overview';
import { QuickSearch, ResourceDialog } from './dialogs';
import { exportBackup, parseBackup, downloadFile } from './backup';
import { DocumentOutline, DocumentOutlineRail, Backlinks } from './document-panels';
import { DocumentOrganizationControls } from './organization-ui';
import { TagBadge } from './tag-ui';
import { Workbench } from './workbench-ui';
import { openLocation, activeTab, closeTab, navigateHistory, changeTabMode, emptyWorkbench, mergePanes, openTab, reconcileWorkbench, restoreWorkbench, type PaneId, type WorkbenchTab } from './workbench';
import { TrashDialog } from './trash-ui';
import './styles.css';
import './workbench-styles.css';

function useChanges(subscribe: (listener: () => void) => () => void) {
  const [, rerender] = useReducer(value => value + 1, 0);
  useEffect(() => subscribe(rerender), [subscribe]);
}
function SaveStatus({ persistence, sync }: { persistence: PersistentDocument;sync?:CloudStorage }) {
  useChanges(persistence.subscribe);
  const [,redraw]=useReducer(value=>value+1,0);useEffect(()=>sync?.subscribe(redraw),[sync]);
  const transport=sync&&(sync.memoryOnly?'SQLite 未保存全部改动，请保留页面或导出草稿':sync.status==='blocked'?sync.reason:sync.status==='offline'?'已断线，改动待提交':sync.status==='connecting'?'正在连接与补齐同步':sync.isReadOnly(persistence.id)?(sync.pendingCount?'只读 · 待提交草稿已保留':'只读 · 已同步'):sync.pendingCount?'改动已暂存 SQLite，等待提交':undefined);
  const message=persistence.state==='saving'?'正在保存…':persistence.state==='error'&&sync?.status==='connected'&&!sync.memoryOnly?'改动待提交，请重试保存':transport??({ loading: persistence.remote?'正在读取共享文档':'正在读取本地数据', saving: '正在保存…', saved: persistence.remote?'已同步到工作空间':'已保存到本机', error: '保存失败，改动仍在内存中' })[persistence.state];
  return <span className={`save-status ${persistence.state}`} role="status">
    <span className="save-status-message" title={message}>{message}</span>
    {sync&&<small> · 已落盘待提交 {sync.pendingCount} 条{sync.lastSuccess?` · 最后确认 ${new Date(sync.lastSuccess).toLocaleTimeString()}`:''}</small>}
    <span className="save-status-actions">{!sync?.readOnly && (persistence.state === 'error'||sync?.status==='offline'||!!sync?.pendingCount) && <button onClick={() => void (async()=>{await sync?.reconnect();if(persistence.state==='error'&&!sync?.isReadOnly(persistence.id))await persistence.retry();})().catch(() => {})}>重试保存</button>}{persistence.remote&&persistence.state==='error'&&<button onClick={()=>downloadFile('文档草稿.mydoc-draft.json',JSON.stringify({format:'mydoc-document-draft',version:1,id:persistence.id,update:Array.from(Y.encodeStateAsUpdate(persistence.doc))}),'application/json')}>导出文档草稿</button>}</span>
  </span>;
}
function DocumentEditor({ item, meta, workspace, onDelete, busy, name, onMove, onNavigate,loadWorkspace,boardRequest, tab, onMode, active, onSaveTemplate, commandRequest, onCommandHandled, onAttachFiles }: {
  onAttachFiles:(files:File[],documentId:string)=>Promise<{id:string;title:string}>;loadWorkspace:(id:string)=>Promise<LocalWorkspace>;onCommandHandled:()=>void; commandRequest?: {kind:'find'|'export'|'markdown';tick:number}; onSaveTemplate: () => void; item: OpenDocument; meta: DocumentMeta; workspace: LocalWorkspace; onDelete: () => void; busy: boolean; name: string; onMove: (folderId: string) => void; onNavigate: (id: string, workspaceId: string, frameId?:string, other?:boolean) => void;boardRequest?:{frameId:string;tick:number}; tab: WorkbenchTab; active: boolean; onMode: (mode: 'page'|'board') => void;
}) {
  const { store } = item;
  const readOnly=workspace.documentReadOnly(meta.id);busy=busy||readOnly;
  const [linkNotice,setLinkNotice]=useState('');const [settingsOpen,setSettingsOpen]=useState(false); const [toolsOpen,setToolsOpen]=useState(false);const [permissionsOpen,setPermissionsOpen]=useState(false);const settings=readEditorSettings(workspace.uiPreferences.get('editorSettings'));
  useChanges(workspace.subscribe);
  const saveSettings=async(value:typeof settings)=>{const previous=workspace.uiPreferences.get('editorSettings');if(workspace.persistence.state==='error')await workspace.persistence.retry();workspace.uiPreferences.set('editorSettings',value);try{await workspace.persistence.flush();}catch(error){if(previous===undefined)workspace.uiPreferences.delete('editorSettings');else workspace.uiPreferences.set('editorSettings',previous);throw error;}};
  const [, refresh] = useReducer(value => value + 1, 0);
  useEffect(() => { store.doc.on('update', refresh); return () => store.doc.off('update', refresh); }, [store]);
  const mode = tab.mode; const setMode = onMode; const titleId = useId(); const [referenceSide, setReferenceSide] = useState('tab'); const navigate = (id: string, ws: string, frame?: string) => { onNavigate(id, ws, frame, referenceSide === 'other'); }; 
  const history = store.history; const [outlineOpen, setOutlineOpen] = useState(false); const [referenceRequest, setReferenceRequest] = useState(0); const [headingRequest, setHeadingRequest] = useState<{ position: number; tick: number }>();
  const outline = documentOutline(readDocument(store), (id, ws) => ws === workspace.id ? workspace.documents.get(id)?.title : undefined);
  return <DocumentPeekScope workspace={workspace} loadWorkspace={loadWorkspace} onNavigate={onNavigate} active={active} mode={mode}>
    {settingsOpen&&<EditorSettingsDialog initial={settings} onSave={saveSettings} onClose={()=>{setSettingsOpen(false);requestAnimationFrame(()=>{if(active)(document.getElementById(titleId)?.closest('article')?.querySelector('.document-editor:not([hidden]) .ProseMirror') as HTMLElement|null)?.focus();});}}/>}{permissionsOpen&&workspace.sync&&<DocumentPermissionDialog workspace={workspace} documentId={meta.id} title={meta.title} onClose={()=>setPermissionsOpen(false)}/>}
    <header className="editor-header"><span>{name} / 文档</span><div className="header-actions"><div className="editor-mode" role="group" aria-label="编辑模式"><button aria-pressed={mode === 'page'} onClick={() => setMode('page')}>页面</button><button aria-pressed={mode === 'board'} onClick={() => { readDocument(store); store.history.stopCapturing(); setMode('board'); }}>白板</button></div><button className="editor-tools-toggle" aria-expanded={toolsOpen} aria-controls={`${titleId}-tools`} onClick={()=>setToolsOpen(value=>!value)}>文档工具</button><div id={`${titleId}-tools`} className={`editor-secondary ${toolsOpen ? 'tools-open' : ''}`}><label className="sr-only" htmlFor={`${titleId}-reference`}>引用打开位置</label><select id={`${titleId}-reference`} aria-label="引用打开位置" value={referenceSide} onChange={event=>setReferenceSide(event.target.value)}><option value="tab">引用打开新标签</option><option value="other">引用在另一侧打开</option></select><button disabled={busy || mode === 'board'} onClick={() => setReferenceRequest(value => value + 1)}>插入引用</button><button disabled={mode === 'board'} aria-expanded={outlineOpen} onClick={() => setOutlineOpen(value => !value)}>大纲</button><FolderSelect workspace={workspace} value={meta.folderId} label="文档所在文件夹" disabled={busy} onChange={onMove} /><DocumentOrganizationControls workspace={workspace} meta={meta} busy={busy} />{window.mydocDesktop&&<button onClick={()=>{void navigator.clipboard.writeText(documentLink({workspace:workspace.id,doc:meta.id,mode,frame:tab.frameId,block:tab.blockId})).then(()=>setLinkNotice('桌面链接已复制')).catch(()=>setLinkNotice('复制失败，请重试'));}}>复制桌面链接</button>}{linkNotice&&<span role="status">{linkNotice}</span>}<button disabled={busy} onClick={onSaveTemplate}>保存为模板</button></div>{workspace.sync?.canManageDocuments&&<button title="管理文档权限" aria-label="管理文档权限" onClick={()=>setPermissionsOpen(true)}><Shield size={17}/></button>}<SaveStatus persistence={item.persistence} sync={workspace.sync} /><button aria-label="撤销" disabled={readOnly||!history.canUndo()} onClick={() => { history.undo(); refresh(); }}><Undo2 size={17} /></button><button aria-label="重做" disabled={readOnly||!history.canRedo()} onClick={() => { history.redo(); refresh(); }}><Redo2 size={17} /></button><button aria-label="删除文档" disabled={busy} onClick={onDelete}><Trash2 size={17} /></button></div></header>
    {mode === 'board' ? <><div className="whiteboard-title"><label className="sr-only" htmlFor={titleId}>文档标题</label><input id={titleId} aria-label="文档标题" value={store.meta.get('title') ?? meta.title} disabled={busy} onChange={event => workspace.rename(meta.id,event.target.value,store)} /></div><Whiteboard readOnly={readOnly} onAttachFiles={files=>onAttachFiles(files,meta.id)} assets={workspace.documentAssets(meta.id)} frameRequest={boardRequest} store={store} documentId={meta.id} workspaceId={workspace.id} documents={workspace.list()} embedSource={workspace.embedSource} onNavigate={navigate} /><Backlinks workspace={workspace} documentId={meta.id} onNavigate={navigate} /></> : <div className="document-layout"><article className="document" style={{'--document-width':settings.width==='full'?'none':settings.width==='wide'?'1280px':'1000px','--document-font-size':`${settings.fontSize}px`,'--document-line-height':settings.lineHeight} as CSSProperties}>
      <label className="sr-only" htmlFor={titleId}>文档标题</label>
      <input id={titleId} className="title-input" disabled={busy} value={store.meta.get('title') ?? meta.title} placeholder="无标题文档" onChange={event => workspace.rename(meta.id, event.target.value, store)} />
      <div className="document-tags" aria-label="文档标签">{(meta.tagIds ?? []).map(id => workspace.organization.tags.get(id)).filter(tag => !!tag).map(tag => <TagBadge key={tag.id} tag={tag}/>)}</div><details className="document-meta"><summary>信息</summary><div>创建于 {new Date(meta.createdAt).toLocaleDateString('zh-CN')}</div>{meta.journalDate&&<div>日记日期：{meta.journalDate}</div>}</details>
      <MarkdownDocumentEditor readOnly={readOnly} blockRequest={tab.blockId?{id:tab.blockId,tick:tab.tick??0}:undefined} settings={settings} onOpenSettings={()=>setSettingsOpen(true)} onCommandHandled={onCommandHandled} commandRequest={commandRequest} focusEnabled={active} store={store} assets={workspace.documentAssets(meta.id)} embedSource={workspace.embedSource} documentId={meta.id} workspaceId={workspace.id} documents={workspace.list()} onNavigate={navigate} referenceRequest={referenceRequest} headingRequest={headingRequest} /><Backlinks workspace={workspace} documentId={meta.id} onNavigate={navigate} />
    </article><DocumentOutlineRail entries={outline} onJump={position => setHeadingRequest({ position, tick: Date.now() })} />{outlineOpen && <DocumentOutline entries={outline} onJump={position => setHeadingRequest({ position, tick: Date.now() })} onClose={() => setOutlineOpen(false)} />}</div>}
  </DocumentPeekScope>;
}
export function App() {
  const desktopUpdate=useDesktopUpdates();
  const appearance=useAppearance(),[themeOpen,setThemeOpen]=useState(false);
  const [cloudOpen,setCloudOpen]=useState(false);
  const [settingsOpen,setSettingsOpen]=useState(false); const [settingsTab,setSettingsTab]=useState<SettingsTab>('local');
  const [registry, setRegistry] = useState<WorkspaceRegistry | null>(null); const [workspace, setWorkspace] = useState<LocalWorkspace | null>(null);
  const [bench, setBench] = useState(emptyWorkbench); const currentTab = activeTab(bench); const selected = currentTab?.docId ?? null;
  const setSelected = (id: string | null) => { if (id) { setOverviewOpen(false); setBench(state => openTab(state, id)); } };
  const boardRequest = currentTab?.frameId ? {docId:currentTab.docId,workspaceId:workspace?.id ?? '',frameId:currentTab.frameId,tick:currentTab.tick ?? 0} : undefined;
  const [error, setError] = useState(''); const [notice, setNotice] = useState(''); const [busy, setBusy] = useState(false); const [retry, setRetry] = useState(0); const [confirmDelete, setConfirmDelete] = useState(false); const [searchOpen, setSearchOpen] = useState(false);
  const [editorCommand,setEditorCommand]=useState<{docId:string;pane:PaneId;kind:'find'|'export'|'markdown';tick:number}>();const commandTick=useRef(0);
  const [clipperOpen,setClipperOpen]=useState(false);
  const [backupOpen,setBackupOpen]=useState(false);
  const [teamRestore,setTeamRestore]=useState<{file:File;archive:CloudWorkspaceArchive}>();
  const [draftDialog,setDraftDialog]=useState<{file?:File}|null>(null);
  const [importBusy,setImportBusy]=useState(false);
  const [importDialog,setImportDialog]=useState<{kind:ImportKind;files?:File[];folderId?:string}|null>(null);
  const [templateDialog,setTemplateDialog]=useState<{saveDocId?:string}|null>(null);
  const [newFile,setNewFile]=useState<{folderId?:string;officeOnly?:boolean}|null>(null);
  const [permissionDoc,setPermissionDoc]=useState<string|null>(null);
  const [trashOpen,setTrashOpen]=useState(false);const [journalDate,setJournalDate]=useState<string|null>(null);const [overviewOpen,setOverviewOpen]=useState(false);
  const [dialog, setDialog] = useState<{ kind: 'create' | 'rename' | 'delete' | 'folder'; folder?: FolderMeta; parentId?: string } | null>(null);
  const [documentAction, setDocumentAction] = useState<{id:string;kind:'rename'|'move'|'delete'}|null>(null);
  const [, refresh] = useReducer(value => value + 1, 0); const markdownInput = useRef<HTMLInputElement>(null); const backupInput = useRef<HTMLInputElement>(null);
  const activate = (ws: LocalWorkspace, requested?: string | null) => { const saved = ws.uiPreferences.get('workbench'); let next = restoreWorkbench(saved, ws.resourceIds()); const params=new URLSearchParams(location.hash.slice(1)); if(requested && ws.resourceIds().has(requested)) next=openTab(next,requested,next.activePane,params.get('doc')===requested?params.get('frame')??undefined:undefined); else if(saved===undefined && ws.list()[0])next=openTab(next,ws.list()[0].id);  setWorkspace(ws); setBench(next); setOverviewOpen(false); setConfirmDelete(false); setSearchOpen(false); setSettingsTab('local');setTrashOpen(false);setJournalDate(null);setTemplateDialog(null);setNewFile(null);setImportDialog(null);setClipperOpen(false);setDocumentAction(null); };
  useLayoutEffect(()=>{if(!workspace)return; const clean=reconcileWorkbench(bench,workspace.resourceIds()); if(clean!==bench){setBench(clean);return;} workspace.uiPreferences.set('workbench',bench);},[workspace,bench,workspace?[...workspace.resourceIds()].join(','):'']);

  useEffect(() => {
    let active = true; const manager = new WorkspaceRegistry(new SQLiteStorage(),new SQLiteAssetStorage());
    void manager.open().then(async () => {if(!window.mydocDesktop||cloudServerUrl())await manager.refreshCloud().catch(error=>{if(active)setNotice('团队工作空间加载失败：'+(error instanceof Error?error.message:'请检查服务器连接后重试'));}); const params = new URLSearchParams(location.hash.slice(1)); const requested = params.get('workspace'); let ws = await manager.select(manager.list().some(entry => entry.id === requested) ? requested! : manager.activeId);
      if (!requested && params.get('doc') && !ws.documents.has(params.get('doc')!) && manager.entries.has('local')) { const legacy = await manager.get('local'); if (legacy.documents.has(params.get('doc')!)) ws = await manager.select('local'); }
      if (active) { setRegistry(manager); activate(ws, params.get('doc')); setError(''); }
    }).catch(error => { if (active) setError(`本地工作区打开失败：${error instanceof Error ? error.message : '存储初始化失败'}。请重试。`); });
    return () => { active = false; void manager.destroy(); };
  }, [retry]);
  useEffect(()=>window.mydocDesktop?.onPrepareClose(async()=>{
    if(busy||importBusy)throw new Error('正在处理导入或其他操作，请完成后再关闭。');
    if(!registry){
      if(error)return;
      throw new Error('工作区尚未打开，请重试打开后再关闭。');
    }
    document.activeElement instanceof HTMLElement&&document.activeElement.blur();
    window.dispatchEvent(new Event('mydoc-before-close'));
    await registry.flush();
    await projectPreferences.flush();
  }),[registry,busy,importBusy]);
  useEffect(() => registry?.subscribe(refresh), [registry]); useEffect(() => workspace?.subscribe(refresh), [workspace]);
  useEffect(() => { const openOverview = () => setOverviewOpen(true); window.addEventListener('open-all-documents', openOverview); return () => window.removeEventListener('open-all-documents', openOverview); }, []);
  useEffect(()=>{const warn=(event:BeforeUnloadEvent)=>{if(workspace?.sync&&workspace.pendingDraft){event.preventDefault();event.returnValue='';}};window.addEventListener('beforeunload',warn);return()=>window.removeEventListener('beforeunload',warn);},[workspace]);
  useEffect(()=>{if(workspace?.readOnly){setImportDialog(null);setClipperOpen(false);setTemplateDialog(null);setDialog(null);setConfirmDelete(false);}},[workspace?.readOnly]);
  const run = async (action: () => Promise<void>) => { setBusy(true); setError(''); setNotice(''); try { await action(); } catch (error) { setError(error instanceof Error ? error.message : '操作失败，请重试'); } finally { setBusy(false); } };
  useEffect(() => { const navigate = () => { if (!registry || !workspace) return; const params = new URLSearchParams(location.hash.slice(1)); const id = params.get('doc'); const ws = params.get('workspace');
      if (ws && ws !== workspace.id && registry.entries.has(ws)) void run(async () => { await workspace.flush(); activate(await registry.select(ws), id); }); else if (id && workspace.resourceIds().has(id)){setBench(state=>openTab(state,id,state.activePane,params.get('frame')??undefined));}
    }; window.addEventListener('hashchange', navigate); return () => window.removeEventListener('hashchange', navigate); }, [registry, workspace]);
  useEffect(() => { if (!workspace) return; const params = new URLSearchParams(location.hash.slice(1)); const frame=boardRequest?.docId===selected&&boardRequest.workspaceId===workspace.id?boardRequest.frameId:null;if (params.get('doc') !== selected || params.get('workspace') !== workspace.id||params.get('frame')!==frame) {if(frame)params.set('frame',frame);else params.delete('frame'); params.set('workspace', workspace.id); if (selected) params.set('doc', selected); else params.delete('doc'); params.delete('block'); history.replaceState(null, '', `#${params}`); } }, [workspace, selected,boardRequest]);
  useEffect(()=>{const shortcut=(event:KeyboardEvent)=>{if(event.defaultPrevented||event.isComposing||busy||!workspace||!((event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==='k'))return;
    const modal=document.querySelector('dialog[open]');if(modal&&modal.getAttribute('aria-label')!=='统一命令面板')return;
    event.preventDefault();if(searchOpen){const input=document.getElementById('document-search') as HTMLInputElement|null;input?.focus();input?.select();}else setSearchOpen(true);
  };window.addEventListener('keydown',shortcut);return()=>window.removeEventListener('keydown',shortcut);},[workspace,busy,searchOpen]);
  useEffect(()=>{if(workspace&&selected&&paletteHistory(workspace.uiPreferences.get('commandHistory'))[0]!=='doc:'+selected)workspace.uiPreferences.set('commandHistory',rememberPalette(workspace.uiPreferences.get('commandHistory'),'doc:'+selected));},[workspace,selected]);

  const docs = workspace?.list() ?? []; const meta = docs.find(doc => doc.id === selected); const workspaceName = registry?.entries.get(workspace?.id ?? '')?.name ?? '本地工作区';
  const todayJournal=async()=>{if(!workspace)return;const date=localJournalDate();if(workspace.recycledJournal(date)){setJournalDate(date);return;}setSelected(await workspace.getOrCreateJournal(date));};
  useDocumentLinkIntake({ready:!!registry&&!!workspace,busy:busy||importBusy,notify:setError,open:async target=>{
    if(!registry||!workspace)return;
    if(!registry.entries.has(target.workspace))throw new Error('链接对应的工作区不在本机，请先导入备份或确认链接来源。');
    const destination=await registry.get(target.workspace);
    if(!destination.list().some(doc=>doc.id===target.doc))throw new Error('链接对应的文档不存在或已在回收站，请检查链接或恢复文档。');
    setBusy(true);try{await workspace.flush();if(destination!==workspace){await registry.select(destination.id);activate(destination);setBench(state=>openLocation(restoreWorkbench(destination.uiPreferences.get('workbench'),new Set(destination.list().map(doc=>doc.id))),target));}else setBench(state=>openLocation(state,target));setError('');}finally{setBusy(false);}
  }});
  const fileDragging=useFileIntake({ready:!!workspace&&!workspace.readOnly,busy:busy||importBusy,onImport:files=>{setError('');setImportDialog({kind:'markdown',files});},notify:setError});
  const createAttachmentDocument=async(files:File[],_sourceDocumentId:string)=>{
    if(!workspace||busy||importBusy)throw new Error('正在处理其他操作，请稍后重试');
    if(!files.length||files.length>10)throw new Error('每次最多添加 10 个附件');
    const ws=workspace;setBusy(true);
    const title=('附件 · '+files[0].name+(files.length>1?` 等 ${files.length} 个文件`:'')).slice(0,160);let id='';
    try{ id=await ws.create(title);const records=[];for(const file of files)records.push(await ws.documentAssets(id).save(file,false));
      const content=documentSchema.nodes.doc.create(null,records.map(asset=>documentSchema.nodes.paragraph.create(null,documentSchema.nodes.attachment.create({assetId:asset.id,name:asset.name,size:asset.size,mime:asset.mime}))));
      const item=await ws.openDocument(id);writeDocument(item.store,content);await item.persistence.flush();return {id,title};
    }catch(error){if(id)await ws.delete(id).catch(()=>{});throw error;}finally{setBusy(false);}
  };
  const create = (folderId?: string) => {if(workspace&&!workspace.readOnly)setNewFile({folderId});};
  const toggleFavorite = (id: string) => void run(async () => { if (!workspace) return; if (workspace.persistence.state === 'error') await workspace.persistence.retry(); const before = !!workspace.fileMeta(id)?.favorite; workspace.setFileFavorite(id, !before); try { await workspace.persistence.flush(); } catch (error) { workspace.setFileFavorite(id, before); throw error; } });
  const selectWorkspace = (id: string) => run(async () => { if (!registry || !workspace) return; await workspace.flush(); activate(await registry.select(id)); });
  const refreshCloudWorkspace = async () => {
    if (!registry) return;
    await registry.flush();
    await registry.switchServer();
    if (!window.mydocDesktop || cloudServerUrl()) await registry.refreshCloud();
    await registry.flush();
    if (workspace && !registry.entries.has(workspace.id)) activate(await registry.select(registry.activeId));
  };
  const remove = () => run(async () => { if (!workspace || !selected) return; setConfirmDelete(false); await workspace.delete(selected); setBench(state=>reconcileWorkbench(state,workspace.resourceIds())); });
  const documentRequired=selected&&workspace?.documents.has(selected)?undefined:'请先打开一篇笔记文档';const pageRequired=documentRequired||(currentTab?.mode==='board'?'请先切换到页面模式':undefined);
  const editorAction=(kind:'find'|'export'|'markdown')=>{if(selected)setEditorCommand({docId:selected,pane:bench.activePane,kind,tick:++commandTick.current});};
  const changeDocumentMode=async(mode:'page'|'board')=>{if(!selected||!workspace)return;const pane=bench.activePane,item=await workspace.openDocument(selected);readDocument(item.store);item.store.history.stopCapturing();setBench(state=>changeTabMode(state,pane,selected,mode));};
  const commands:PaletteCommand[]=workspace?[
    {id:'navigation-history',label:'打开导航历史',keywords:'history navigation 最近访问',disabled:docs.length?undefined:'请先创建文档',run:()=>{window.dispatchEvent(new CustomEvent('open-document-navigation',{detail:workspace.id}));}},
    {id:'navigation-back',label:'导航后退',keywords:'back previous 返回',disabled:(bench[bench.activePane].history?.cursor??-1)>0?undefined:'没有更早的访问记录',run:()=>setBench(state=>navigateHistory(state,state.activePane,(state[state.activePane].history?.cursor??0)-1))},
    {id:'navigation-forward',label:'导航前进',keywords:'forward next',disabled:(bench[bench.activePane].history?.cursor??-1)>=0&&(bench[bench.activePane].history?.cursor??0)<(bench[bench.activePane].history?.entries.length??0)-1?undefined:'没有后续访问记录',run:()=>setBench(state=>navigateHistory(state,state.activePane,(state[state.activePane].history?.cursor??0)+1))},
    {id:'new-document',label:'新建文档',keywords:'create new document 创建',run:()=>setNewFile({})},
    {id:'today-journal',label:'打开今日日记',keywords:'journal diary daily 日历',run:todayJournal},
    {id:'journal-calendar',label:'打开日历笔记',keywords:'journal calendar 日期日记',run:()=>setJournalDate(localJournalDate())},
    {id:'templates',label:'从模板新建',keywords:'template 模板文档',run:()=>setTemplateDialog({})},
    {id:'save-template',label:'保存为文档模板',keywords:'save template',disabled:documentRequired,run:()=>{if(selected)setTemplateDialog({saveDocId:selected});}},
    {id:'document-permissions',label:'管理文档权限',keywords:'permission access 文档权限 授权',disabled:documentRequired||!workspace?.sync?.canManageDocuments?'需要团队工作空间管理员权限':undefined,run:()=>{if(selected)setPermissionDoc(selected);}},
    {id:'page-mode',label:'切换到页面',keywords:'page document',disabled:documentRequired||(currentTab?.mode==='page'?'当前已是页面模式':undefined),run:()=>changeDocumentMode('page')},
    {id:'board-mode',label:'切换到白板',keywords:'board canvas',disabled:documentRequired||(currentTab?.mode==='board'?'当前已是白板模式':undefined),run:()=>changeDocumentMode('board')},
    {id:'open-other',label:'在另一侧打开文档',keywords:'split 分屏 左右',disabled:documentRequired,run:()=>{if(selected)setBench(state=>openTab(state,selected,state.activePane==='left'?'right':'left'));}},
    {id:'merge-panes',label:'合并分屏',keywords:'merge split 单栏',disabled:bench.split?undefined:'当前没有分屏',run:()=>setBench(mergePanes)},
    {id:'favorite',label:meta?.favorite?'取消收藏文档':'收藏文档',keywords:'favorite star',disabled:documentRequired,run:async()=>{if(selected){if(workspace.persistence.state==='error')await workspace.persistence.retry();const before=!!workspace.documents.get(selected)?.favorite;workspace.organization.setFavorite(selected,!before);try{await workspace.persistence.flush();}catch(error){workspace.organization.setFavorite(selected,before);throw error;}}}},
    {id:'find',label:'页内查找',keywords:'find 搜索正文',disabled:pageRequired,run:()=>editorAction('find')},
    {id:'export',label:'导出 PDF / HTML',keywords:'export print 打印',disabled:pageRequired,run:()=>editorAction('export')},
    {id:'markdown',label:'导出 Markdown',keywords:'export md',disabled:pageRequired,run:()=>editorAction('markdown')},
    {id:'clipper',label:'网页剪藏',keywords:'import-clipper url web 网页 收藏',run:()=>setClipperOpen(true)},
    {id:'import',label:'导入 Markdown / Notion',keywords:'import zip 迁移',run:()=>setImportDialog({kind:'markdown'})},
    ...(window.mydocDesktop?[{id:'desktop-update',label:'版本与更新',keywords:'desktop update version 检查更新',run:()=>desktopUpdate.setOpen(true)}]:[]),
    {id:'database-backup',label:'自动备份与恢复',keywords:'backup restore sqlite 数据保护',run:()=>setBackupOpen(true)},
    {id:'draft-restore',label:'从草稿恢复',keywords:'draft recover offline 草稿 离线 未提交',run:()=>setDraftDialog({})},
    {id:'trash',label:'打开回收站',keywords:'trash restore 恢复文档',run:()=>setTrashOpen(true)},
    {id:'organization',label:'管理标签与集合',keywords:'tags collection organize 设置',disabled:workspace.readOnly?'当前工作区为只读':undefined,run:()=>{setSettingsTab('organization');setSettingsOpen(true);}},
  ]:[];
  return <div className="app-shell">{settingsOpen&&<SettingsDialog key={settingsTab} initialTab={settingsTab} workspace={workspace} activeWorkspaceId={workspace?.id} workspaces={(registry?.list()??[]).filter(entry=>entry.cloud).map(entry=>({...entry,role:entry.role??'member'}))} onClose={()=>setSettingsOpen(false)} theme={appearance.mode} onThemeSave={appearance.save} onRefresh={async()=>{await registry?.flush();}} onCloudSessionChanged={refreshCloudWorkspace}/>} {cloudOpen&&<CloudDialog onPrepareSwitch={async()=>{await registry?.flush();await projectPreferences.flush();}} onClose={()=>setCloudOpen(false)} onRefresh={refreshCloudWorkspace} onOpen={id=>{setCloudOpen(false);void selectWorkspace(id);}}/>}<DocumentSidebar onCreateOffice={()=>setNewFile({officeOnly:true})} onSettings={()=>{setSettingsTab('local');setSettingsOpen(true)}} onCloud={()=>setCloudOpen(true)} onUpdates={()=>desktopUpdate.setOpen(true)} updateNotice={desktopUpdate.state?.phase==='downloaded'?'待安装':desktopUpdate.state?.phase==='available'?'有更新':undefined} onDatabaseBackup={()=>setBackupOpen(true)} onClipper={()=>setClipperOpen(true)} onToday={()=>void run(todayJournal)} onJournal={()=>setJournalDate(localJournalDate())} workspace={workspace} workspaces={registry?.list() ?? []} selected={selected} busy={busy || importBusy} onCreate={folderId => void create(folderId)} onSelect={id => { setSelected(id); setConfirmDelete(false); }} onWorkspace={id => void selectWorkspace(id)} onSearch={() => setSearchOpen(true)} onWorkspaceAction={kind => setDialog({ kind })} onTemplates={()=>setTemplateDialog({})} onTrash={()=>setTrashOpen(true)} onDocumentAction={(id,kind)=>setDocumentAction({id,kind})} onFolder={(folder,parentId) => setDialog({ kind: 'folder', folder, parentId })} onImport={folderId=>setImportDialog({kind:'markdown',folderId})} onNotionImport={folderId=>setImportDialog({kind:'notion',folderId})} onRestore={() => backupInput.current?.click()} onDraftRestore={()=>setDraftDialog({})} onBackup={() => void run(async () => { if (!workspace) return; if(workspace.sync){await downloadTeamWorkspaceBackup(workspace,workspaceName);setNotice('团队工作区完整备份已下载');return;}const data = await exportBackup(workspace, workspaceName); downloadFile(`${workspaceName.replace(/[\\/:*?"<>|]/g, '_')}.mydoc.json`, JSON.stringify(data), 'application/json'); setNotice('工作区备份已下载'); })} />
    <input ref={markdownInput} type="file" hidden multiple accept=".md,.markdown,.txt,text/markdown,text/plain" aria-label="导入 Markdown 文件" onChange={event => { const files = [...(event.target.files ?? [])]; event.target.value = ''; if(files.length)setImportDialog({kind:'markdown',files}); }} />
    <input ref={backupInput} type="file" hidden accept=".json,application/json" aria-label="恢复工作区备份文件" onChange={event => { const file = event.target.files?.[0]; event.target.value = ''; if (file) void run(async () => { if (!registry || !workspace) return; if (file.size > 50 * 1024 * 1024) throw new Error('备份超过 50 MB，请分工作区备份'); const source=await file.text(),input:unknown=JSON.parse(source);if(isCloudWorkspaceArchive(input)){setTeamRestore({file,archive:input});return;}const backup = parseBackup(source); await workspace.flush(); activate(await registry.restore(backup)); setNotice(`已恢复 ${backup.documents.length} 篇文档和 ${backup.trash?.length??0} 篇回收文档、${backup.office?.office.filter(r=>!r.deleted).length??0} 个办公文件到新工作区`); }); }} />
    <main>{fileDragging&&<div className="file-drop-hint" role="status">松开以导入文档或添加文件</div>}{error && <div className="error-banner" role="alert">{error}{!workspace && <button onClick={() => setRetry(value => value + 1)}>重新打开</button>}</div>}{notice && <div className="operation-notice" role="status">{notice}<button aria-label="关闭提示" onClick={() => setNotice('')}>×</button></div>}{workspace && <div className="workspace-status"><SaveStatus persistence={workspace.persistence} sync={workspace.sync} />{registry?.persistence.state === 'error' && <SaveStatus persistence={registry.persistence} />}{busy && <span role="status"> · 正在处理…</span>}{workspace.sync&&(workspace.readOnly||workspace.sync.status==='offline'||workspace.persistence.state==='error')&&<div className="sync-protection" role="status"><span>{workspace.sync.needsLogin?'登录已过期。页面内容与已暂存的草稿仍保留，重新登录后会继续提交。':workspace.sync.status==='blocked'?'编辑已暂停。当前页面中的内容仍保留，请先导出草稿。':workspace.readOnly?'当前为查看者权限，可浏览与导出文档。':'尚有改动等待同步，关闭页面前请重试或导出草稿。'}</span><button onClick={()=>void workspace.exportDraft().then(draft=>downloadFile('工作空间草稿.mydoc-draft.json',JSON.stringify(draft),'application/json'))}>导出草稿</button>{workspace.sync.needsLogin&&<button className="primary" onClick={()=>setCloudOpen(true)}>重新登录</button>}{workspace.sync.status==='blocked'&&<button onClick={()=>void workspace.sync?.reconnect()}>{workspace.sync.needsLogin?'登录后继续提交':'重新检查权限'}</button>}</div>}</div>}
      {!workspace && !error ? <div className="empty-state" role="status">正在打开本地工作区…</div> : workspace ? <>{(!!docs.length || !!workspace.officeResources.size) && <Workbench workspace={workspace} state={bench} onChange={setBench} onCloseTab={(pane,id)=>void run(async()=>{await workspace.flushResource(id);setBench(state=>closeTab(state,pane,id));})} busy={busy} render={(opened, tab, pane)=>{ const entry=workspace.documents.get(tab.docId); if(!entry)return null;return <DocumentEditor onAttachFiles={createAttachmentDocument} loadWorkspace={id=>registry!.get(id)} key={`${workspace.id}:${pane}:${tab.docId}`} item={opened} meta={entry} workspace={workspace} name={workspaceName} active={bench.activePane===pane} tab={tab} onMode={mode=>setBench(state=>changeTabMode(state,pane,tab.docId,mode))} boardRequest={tab.frameId?{frameId:tab.frameId,tick:tab.tick??0}:undefined} onNavigate={(id, targetWorkspace, frameId, other)=>{ const targetPane:PaneId=other?(pane==='left'?'right':'left'):pane; if(targetWorkspace===workspace.id){if(!workspace.documents.has(id)){setError('引用的文档已删除，正文内容已保留。');return;}setBench(state=>openTab(state,id,targetPane,frameId));setConfirmDelete(false);}else void run(async()=>{if(!registry||!registry.entries.has(targetWorkspace))throw new Error('引用的工作区不存在，正文内容已保留。');const target=await registry.get(targetWorkspace);if(!target.documents.has(id))throw new Error('引用的文档已删除，正文内容已保留。');await workspace.flush();activate(await registry.select(targetWorkspace),id);});}} onMove={folder=>void run(async()=>{workspace.moveDocument(entry.id,folder||null);await workspace.persistence.flush();})} busy={busy} onCommandHandled={()=>setEditorCommand(undefined)} commandRequest={editorCommand?.docId===entry.id&&editorCommand.pane===pane?editorCommand:undefined} onSaveTemplate={()=>setTemplateDialog({saveDocId:entry.id})} onDelete={()=>setConfirmDelete(true)}/>;}} />}{!docs.length && !workspace.officeResources.size && busy && <div className="empty-state" role="status">正在保存变更…</div>}{!docs.length && !workspace.officeResources.size && !busy && <div className="empty-state"><FileText size={32}/><h1>开始你的第一篇文档</h1><p>{registry?.entries.get(workspace.id)?.cloud?'文档会同步到工作空间，与成员共同编辑。':'记录想法与资料，内容自动保存到本机。'}</p><button className="primary" disabled={busy} onClick={()=>void create()}>新建文档</button></div>}</> : null}
      {confirmDelete && <div className="delete-confirm" role="alert"><span>删除“{meta?.title || '无标题文档'}”？将移入回收站，可随时恢复。</span><button onClick={() => setConfirmDelete(false)}>取消</button><button className="danger" onClick={() => void remove()}>确认删除</button></div>}
    </main>
    {overviewOpen && workspace && <div className="document-overview-layer"><DocumentOverview workspace={workspace} busy={busy || importBusy} onCreate={() => void create()} onSelect={id => setSelected(id)} onFavorite={toggleFavorite} onAction={(id, kind) => setDocumentAction({ id, kind })} /></div>}
    {clipperOpen&&workspace&&<WebClipperDialog key={workspace.id} workspace={workspace} onClose={()=>setClipperOpen(false)} onFinish={id=>{setClipperOpen(false);setSelected(id);}} onTrash={()=>{setClipperOpen(false);setTrashOpen(true);}}/>}
    {importDialog&&workspace&&<DocumentImportDialog onBusyChange={setImportBusy} key={workspace.id} workspace={workspace} initialKind={importDialog.kind} initialFiles={importDialog.files} initialFolderId={importDialog.folderId} onClose={()=>setImportDialog(null)} onFinish={id=>{setImportDialog(null);if(id)setSelected(id);}}/>}
    {newFile && workspace && <NewFileDialog key={workspace.id} workspace={workspace} folderId={newFile.folderId} officeOnly={newFile.officeOnly} onClose={()=>setNewFile(null)} onCreated={id=>{setNewFile(null);setSelected(id);}}/>}
    {templateDialog && workspace && <DocumentTemplateDialog key={workspace.id} workspace={workspace} saveDocId={templateDialog.saveDocId} onClose={()=>setTemplateDialog(null)} onCreate={id=>{setTemplateDialog(null);setSelected(id);setNotice('已从模板创建文档');}} />}
    {permissionDoc && workspace?.sync && workspace.documents.has(permissionDoc) && <DocumentPermissionDialog workspace={workspace} documentId={permissionDoc} title={workspace.documents.get(permissionDoc)?.title??'文档'} onClose={()=>setPermissionDoc(null)} />}
    {journalDate&&workspace&&<JournalDialog key={workspace.id} workspace={workspace} initialDate={journalDate} onClose={()=>setJournalDate(null)} onOpen={id=>{setJournalDate(null);setSelected(id);setConfirmDelete(false);}}/>}
    {trashOpen && workspace && <TrashDialog key={workspace.id} workspace={workspace} onClose={()=>setTrashOpen(false)} onRestore={id=>{setTrashOpen(false);setSelected(id);setNotice('文档已恢复，引用将重新生效');}} />}
    {themeOpen&&<ThemeDialog initial={appearance.mode} preview={appearance.setMode} save={appearance.save} onClose={()=>setThemeOpen(false)}/>}
    {desktopUpdate.open&&<DesktopUpdateDialog state={desktopUpdate.state} onState={desktopUpdate.setState} onClose={()=>desktopUpdate.setOpen(false)}/>}
    {draftDialog&&workspace&&<DraftRestoreDialog key={workspace.id} workspace={workspace} initialFile={draftDialog.file} onBusyChange={setBusy} onClose={()=>setDraftDialog(null)} onFinish={id=>{setDraftDialog(null);if(id){setSelected(id);setConfirmDelete(false);}setNotice('草稿已恢复为新文档');}}/>}
    {teamRestore&&registry&&<TeamWorkspaceRestoreDialog file={teamRestore.file} archive={teamRestore.archive} onClose={()=>setTeamRestore(undefined)} onRestore={async id=>{await registry.refreshCloud();activate(await registry.select(id));setTeamRestore(undefined);setNotice('已恢复到新的团队工作区');}}/>}
    {backupOpen&&registry&&<DatabaseBackupDialog onClose={()=>setBackupOpen(false)} onBusyChange={setBusy} prepare={async()=>{document.activeElement instanceof HTMLElement&&document.activeElement.blur();window.dispatchEvent(new Event('mydoc-before-close'));await registry.flush();await projectPreferences.flush();}}/>}
    {searchOpen && workspace && <QuickSearch commands={commands} workspace={workspace} onClose={() => setSearchOpen(false)} onSelect={id => { setSelected(id); setConfirmDelete(false); }} />}
    {documentAction && workspace && workspace.fileMeta(documentAction.id) && <ResourceDialog key={`${documentAction.kind}:${documentAction.id}`} title={documentAction.kind==='rename'?'重命名文档':documentAction.kind==='move'?'移动文档':'删除文档'} initialName={workspace.fileMeta(documentAction.id)?.title} nameLabel="文档名称" confirm={documentAction.kind==='delete'} hideName={documentAction.kind==='move'} destructive={documentAction.kind==='delete'} onClose={()=>setDocumentAction(null)} action={async(name,form)=>{
      const id=documentAction.id;
      if(documentAction.kind==='delete'){if(workspace.officeResources.has(id))await workspace.deleteOfficeResource(id);else await workspace.delete(id);setBench(state=>reconcileWorkbench(state,workspace.resourceIds()));}
      else if(documentAction.kind==='move'){workspace.moveDocument(id,String(form.get('folder')||'')||null);await workspace.persistence.flush();}
      else if(workspace.officeResources.has(id)){await workspace.renameOfficeResource(id,name.trim());}
      else {if(!name.trim())throw new Error('请输入文档名称');const item=await workspace.openDocument(id);workspace.rename(id,name.trim(),item.store);await item.persistence.flush();await workspace.persistence.flush();}
    }}>
      {documentAction.kind==='move'&&<label className="form-field">目标文件夹<FolderSelect workspace={workspace} name="folder" label="目标文件夹" value={workspace.fileMeta(documentAction.id)?.folderId}/></label>}
      {documentAction.kind==='delete'&&<p>删除“{workspace.fileMeta(documentAction.id)?.title||'无标题文档'}”？{workspace.officeResources.has(documentAction.id)?'请先保存命名版本；删除后可从已删除文件版本中恢复。':'将移入回收站，可随时恢复。'}</p>}
    </ResourceDialog>}
    {dialog && registry && workspace && <ResourceDialog key={`${dialog.kind}:${dialog.folder?.id ?? ''}`} title={dialog.kind === 'folder' ? dialog.folder ? '管理文件夹' : '新建文件夹' : dialog.kind === 'create' ? '新建工作区' : dialog.kind === 'rename' ? '重命名工作区' : '删除工作区'} initialName={dialog.kind === 'folder' ? dialog.folder?.name : dialog.kind === 'rename' ? workspaceName : ''} nameLabel={dialog.kind === 'folder' ? '文件夹名称' : '工作区名称'} confirm={dialog.kind === 'delete'} destructive={dialog.kind === 'delete'} onClose={() => setDialog(null)} action={async (name, form) => {
      if (dialog.kind === 'create') { await workspace.flush(); activate(await registry.create(name)); }
      else if (dialog.kind === 'rename') await registry.rename(workspace.id, name);
      else if (dialog.kind === 'delete') {  try { await registry.delete(workspace.id); activate(await registry.select(registry.activeId)); } catch (error) { if (!registry.entries.has(workspace.id)) activate(await registry.select(registry.activeId)); throw error; } }
      else { const parent = String(form.get('parent') || '') || null; if (dialog.folder) { workspace.moveFolder(dialog.folder.id, parent); workspace.renameFolder(dialog.folder.id, name); } else workspace.createFolder(name, parent); await workspace.persistence.flush(); }
    }}>
      {dialog.kind === 'delete' && <p>删除“{workspaceName}”及其中的 {docs.length} 篇文档及 {workspace.trash.size} 篇回收文档？删除后无法恢复，建议先备份。</p>}
      {dialog.kind === 'folder' && <><label className="form-field">上级文件夹<FolderSelect workspace={workspace} name="parent" label="上级文件夹" value={dialog.folder?.parentId ?? dialog.parentId} exclude={dialog.folder?.id} /></label>{dialog.folder && <button type="button" className="folder-delete" onClick={() => void run(async () => { workspace.deleteFolder(dialog.folder!.id); await workspace.persistence.flush(); setDialog(null); })}>删除文件夹，保留文档</button>}</>}
    </ResourceDialog>}
  </div>;
}
