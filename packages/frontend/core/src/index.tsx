import {DocumentPeekScope} from './peek-view-ui';
import {WebClipperDialog} from './web-clipper-ui';
import {JournalDialog} from './journal-ui';
import {localJournalDate} from './journal';
export {initializeTheme} from './theme';
import {useAppearance,ThemeDialog} from './theme-ui';
import {EditorSettingsDialog} from './editor-settings-ui';
import { DocumentImportDialog } from './document-import-ui';
import type { ImportKind } from './import-source';
import { rememberPalette, paletteHistory, type PaletteCommand } from './command-palette';
import { DocumentTemplateDialog } from './document-template-ui';
import { useEffect, useReducer, useRef, useState, useId, useLayoutEffect, type CSSProperties } from 'react';
import { FileText, Trash2, Undo2, Redo2 } from 'lucide-react';
import { SQLiteStorage, SQLiteAssetStorage, type PersistentDocument } from '@mydoc/nbstore';
import { DocumentEditor as MarkdownDocumentEditor, readEditorSettings, readDocument, documentOutline, Whiteboard } from '@mydoc/editor-blocks';
import { LocalWorkspace, type DocumentMeta, type OpenDocument, type FolderMeta } from './workspace';
import { WorkspaceRegistry } from './workspaces';
import { DocumentSidebar, FolderSelect } from './sidebar';
import { QuickSearch, ResourceDialog } from './dialogs';
import { exportBackup, parseBackup, downloadFile } from './backup';
import { DocumentOutline, Backlinks } from './document-panels';
import { OrganizationManager, DocumentOrganizationControls } from './organization-ui';
import { Workbench } from './workbench-ui';
import { activeTab, closeTab, navigateHistory, changeTabMode, emptyWorkbench, mergePanes, openTab, reconcileWorkbench, restoreWorkbench, type PaneId, type WorkbenchTab } from './workbench';
import { TrashDialog } from './trash-ui';
import './styles.css';

function useChanges(subscribe: (listener: () => void) => () => void) {
  const [, rerender] = useReducer(value => value + 1, 0);
  useEffect(() => subscribe(rerender), [subscribe]);
}
function SaveStatus({ persistence }: { persistence: PersistentDocument }) {
  useChanges(persistence.subscribe);
  return <span className={`save-status ${persistence.state}`} role="status">
    {({ loading: '正在读取本地数据', saving: '正在保存…', saved: '已保存到本机', error: '保存失败，改动仍在内存中' })[persistence.state]}
    {persistence.state === 'error' && <button onClick={() => void persistence.retry().catch(() => {})}>重试保存</button>}
  </span>;
}
function DocumentEditor({ item, meta, workspace, onDelete, busy, name, onMove, onNavigate,loadWorkspace,boardRequest, tab, onMode, active, onSaveTemplate, commandRequest, onCommandHandled }: {
  loadWorkspace:(id:string)=>Promise<LocalWorkspace>;onCommandHandled:()=>void; commandRequest?: {kind:'find'|'export'|'markdown';tick:number}; onSaveTemplate: () => void; item: OpenDocument; meta: DocumentMeta; workspace: LocalWorkspace; onDelete: () => void; busy: boolean; name: string; onMove: (folderId: string) => void; onNavigate: (id: string, workspaceId: string, frameId?:string, other?:boolean) => void;boardRequest?:{frameId:string;tick:number}; tab: WorkbenchTab; active: boolean; onMode: (mode: 'page'|'board') => void;
}) {
  const { store } = item;
  const [settingsOpen,setSettingsOpen]=useState(false);const settings=readEditorSettings(workspace.uiPreferences.get('editorSettings'));
  useChanges(workspace.subscribe);
  const saveSettings=async(value:typeof settings)=>{const previous=workspace.uiPreferences.get('editorSettings');if(workspace.persistence.state==='error')await workspace.persistence.retry();workspace.uiPreferences.set('editorSettings',value);try{await workspace.persistence.flush();}catch(error){if(previous===undefined)workspace.uiPreferences.delete('editorSettings');else workspace.uiPreferences.set('editorSettings',previous);throw error;}};
  const [, refresh] = useReducer(value => value + 1, 0);
  useEffect(() => { store.doc.on('update', refresh); return () => store.doc.off('update', refresh); }, [store]);
  const mode = tab.mode; const setMode = onMode; const titleId = useId(); const [referenceSide, setReferenceSide] = useState('tab'); const navigate = (id: string, ws: string, frame?: string) => { onNavigate(id, ws, frame, referenceSide === 'other'); }; 
  const history = store.history; const [outlineOpen, setOutlineOpen] = useState(false); const [referenceRequest, setReferenceRequest] = useState(0); const [headingRequest, setHeadingRequest] = useState<{ position: number; tick: number }>();
  const outline = documentOutline(readDocument(store), (id, ws) => ws === workspace.id ? workspace.documents.get(id)?.title : undefined);
  return <DocumentPeekScope workspace={workspace} loadWorkspace={loadWorkspace} onNavigate={onNavigate} active={active} mode={mode}>
    {settingsOpen&&<EditorSettingsDialog initial={settings} onSave={saveSettings} onClose={()=>{setSettingsOpen(false);requestAnimationFrame(()=>{if(active)(document.getElementById(titleId)?.closest('article')?.querySelector('.document-editor:not([hidden]) .ProseMirror') as HTMLElement|null)?.focus();});}}/>}
    <header className="editor-header"><span>{name} / 文档</span><div className="header-actions"><div className="editor-mode" role="group" aria-label="编辑模式"><button aria-pressed={mode === 'page'} onClick={() => setMode('page')}>页面</button><button aria-pressed={mode === 'board'} onClick={() => { readDocument(store); store.history.stopCapturing(); setMode('board'); }}>白板</button></div><label className="sr-only" htmlFor={`${titleId}-reference`}>引用打开位置</label><select id={`${titleId}-reference`} aria-label="引用打开位置" value={referenceSide} onChange={event=>setReferenceSide(event.target.value)}><option value="tab">引用打开新标签</option><option value="other">引用在另一侧打开</option></select><button disabled={busy || mode === 'board'} onClick={() => setReferenceRequest(value => value + 1)}>插入引用</button><button disabled={mode === 'board'} aria-expanded={outlineOpen} onClick={() => setOutlineOpen(value => !value)}>大纲</button><FolderSelect workspace={workspace} value={meta.folderId} label="文档所在文件夹" disabled={busy} onChange={onMove} /><DocumentOrganizationControls workspace={workspace} meta={meta} busy={busy} /><button disabled={busy} onClick={onSaveTemplate}>保存为模板</button><SaveStatus persistence={item.persistence} /><button aria-label="撤销" disabled={!history.canUndo()} onClick={() => { history.undo(); refresh(); }}><Undo2 size={17} /></button><button aria-label="重做" disabled={!history.canRedo()} onClick={() => { history.redo(); refresh(); }}><Redo2 size={17} /></button><button aria-label="删除文档" disabled={busy} onClick={onDelete}><Trash2 size={17} /></button></div></header>
    {mode === 'board' ? <><div className="whiteboard-title"><label className="sr-only" htmlFor={titleId}>文档标题</label><input id={titleId} aria-label="文档标题" value={store.meta.get('title') ?? meta.title} disabled={busy} onChange={event => workspace.rename(meta.id,event.target.value,store)} /></div><Whiteboard assets={workspace.assets} frameRequest={boardRequest} store={store} documentId={meta.id} workspaceId={workspace.id} documents={workspace.list()} embedSource={workspace.embedSource} onNavigate={navigate} /><Backlinks workspace={workspace} documentId={meta.id} onNavigate={navigate} /></> : <div className="document-layout"><article className="document" style={{'--document-width':settings.width==='full'?'none':settings.width==='wide'?'1280px':'1000px','--document-font-size':`${settings.fontSize}px`,'--document-line-height':settings.lineHeight} as CSSProperties}>
      <label className="sr-only" htmlFor={titleId}>文档标题</label>
      <input id={titleId} className="title-input" disabled={busy} value={store.meta.get('title') ?? meta.title} placeholder="无标题文档" onChange={event => workspace.rename(meta.id, event.target.value, store)} />
      <div className="document-tags" aria-label="文档标签">{(meta.tagIds ?? []).map(id => workspace.organization.tags.get(id)).filter(tag => !!tag).map(tag => <span key={tag.id}>{tag.name}</span>)}</div><details className="document-meta"><summary>信息</summary><div>创建于 {new Date(meta.createdAt).toLocaleDateString('zh-CN')}</div>{meta.journalDate&&<div>日记日期：{meta.journalDate}</div>}</details>
      <MarkdownDocumentEditor settings={settings} onOpenSettings={()=>setSettingsOpen(true)} onCommandHandled={onCommandHandled} commandRequest={commandRequest} focusEnabled={active} store={store} assets={workspace.assets} embedSource={workspace.embedSource} documentId={meta.id} workspaceId={workspace.id} documents={workspace.list()} onNavigate={navigate} referenceRequest={referenceRequest} headingRequest={headingRequest} /><Backlinks workspace={workspace} documentId={meta.id} onNavigate={navigate} />
    </article>{outlineOpen && <DocumentOutline entries={outline} onJump={position => setHeadingRequest({ position, tick: Date.now() })} onClose={() => setOutlineOpen(false)} />}</div>}
  </DocumentPeekScope>;
}
export function App() {
  const appearance=useAppearance(),[themeOpen,setThemeOpen]=useState(false);
  const [registry, setRegistry] = useState<WorkspaceRegistry | null>(null); const [workspace, setWorkspace] = useState<LocalWorkspace | null>(null);
  const [bench, setBench] = useState(emptyWorkbench); const currentTab = activeTab(bench); const selected = currentTab?.docId ?? null;
  const setSelected = (id: string | null) => { if (id) setBench(state => openTab(state, id)); };
  const boardRequest = currentTab?.frameId ? {docId:currentTab.docId,workspaceId:workspace?.id ?? '',frameId:currentTab.frameId,tick:currentTab.tick ?? 0} : undefined;
  const [error, setError] = useState(''); const [notice, setNotice] = useState(''); const [busy, setBusy] = useState(false); const [retry, setRetry] = useState(0); const [confirmDelete, setConfirmDelete] = useState(false); const [searchOpen, setSearchOpen] = useState(false);
  const [editorCommand,setEditorCommand]=useState<{docId:string;pane:PaneId;kind:'find'|'export'|'markdown';tick:number}>();const commandTick=useRef(0);
  const [clipperOpen,setClipperOpen]=useState(false);
  const [importDialog,setImportDialog]=useState<{kind:ImportKind;files?:File[]}|null>(null);
  const [templateDialog,setTemplateDialog]=useState<{saveDocId?:string}|null>(null);
  const [trashOpen,setTrashOpen]=useState(false);const [journalDate,setJournalDate]=useState<string|null>(null);
  const [organizationOpen, setOrganizationOpen] = useState(false);
  const [dialog, setDialog] = useState<{ kind: 'create' | 'rename' | 'delete' | 'folder'; folder?: FolderMeta } | null>(null);
  const [, refresh] = useReducer(value => value + 1, 0); const markdownInput = useRef<HTMLInputElement>(null); const backupInput = useRef<HTMLInputElement>(null);
  const activate = (ws: LocalWorkspace, requested?: string | null) => { const saved = ws.uiPreferences.get('workbench'); let next = restoreWorkbench(saved, new Set(ws.list().map(doc=>doc.id))); const params=new URLSearchParams(location.hash.slice(1)); if(requested && ws.documents.has(requested)) next=openTab(next,requested,next.activePane,params.get('doc')===requested?params.get('frame')??undefined:undefined); else if(saved===undefined && ws.list()[0])next=openTab(next,ws.list()[0].id);  setWorkspace(ws); setBench(next); setConfirmDelete(false); setSearchOpen(false); setOrganizationOpen(false);setTrashOpen(false);setJournalDate(null);setTemplateDialog(null);setImportDialog(null);setClipperOpen(false); };
  useLayoutEffect(()=>{if(!workspace)return; const clean=reconcileWorkbench(bench,new Set(workspace.list().map(doc=>doc.id))); if(clean!==bench){setBench(clean);return;} workspace.uiPreferences.set('workbench',bench);},[workspace,bench,workspace?.list().map(doc=>doc.id).join(',')]);

  useEffect(() => {
    let active = true; const manager = new WorkspaceRegistry(new SQLiteStorage(),new SQLiteAssetStorage());
    void manager.open().then(async () => { const params = new URLSearchParams(location.hash.slice(1)); const requested = params.get('workspace'); let ws = await manager.select(manager.list().some(entry => entry.id === requested) ? requested! : manager.activeId);
      if (!requested && params.get('doc') && !ws.documents.has(params.get('doc')!) && manager.entries.has('local')) { const legacy = await manager.get('local'); if (legacy.documents.has(params.get('doc')!)) ws = await manager.select('local'); }
      if (active) { setRegistry(manager); activate(ws, params.get('doc')); setError(''); }
    }).catch(() => { if (active) setError('无法打开本地数据库。请检查浏览器是否允许存储数据。'); });
    return () => { active = false; void manager.destroy(); };
  }, [retry]);
  useEffect(() => registry?.subscribe(refresh), [registry]); useEffect(() => workspace?.subscribe(refresh), [workspace]);
  const run = async (action: () => Promise<void>) => { setBusy(true); setError(''); setNotice(''); try { await action(); } catch (error) { setError(error instanceof Error ? error.message : '操作失败，请重试'); } finally { setBusy(false); } };
  useEffect(() => { const navigate = () => { if (!registry || !workspace) return; const params = new URLSearchParams(location.hash.slice(1)); const id = params.get('doc'); const ws = params.get('workspace');
      if (ws && ws !== workspace.id && registry.entries.has(ws)) void run(async () => { await workspace.flush(); activate(await registry.select(ws), id); }); else if (id && workspace.documents.has(id)){setBench(state=>openTab(state,id,state.activePane,params.get('frame')??undefined));}
    }; window.addEventListener('hashchange', navigate); return () => window.removeEventListener('hashchange', navigate); }, [registry, workspace]);
  useEffect(() => { if (!workspace) return; const params = new URLSearchParams(location.hash.slice(1)); const frame=boardRequest?.docId===selected&&boardRequest.workspaceId===workspace.id?boardRequest.frameId:null;if (params.get('doc') !== selected || params.get('workspace') !== workspace.id||params.get('frame')!==frame) {if(frame)params.set('frame',frame);else params.delete('frame'); params.set('workspace', workspace.id); if (selected) params.set('doc', selected); else params.delete('doc'); params.delete('block'); history.replaceState(null, '', `#${params}`); } }, [workspace, selected,boardRequest]);
  useEffect(()=>{const shortcut=(event:KeyboardEvent)=>{if(event.defaultPrevented||event.isComposing||busy||!workspace||!((event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==='k'))return;
    const modal=document.querySelector('dialog[open]');if(modal&&modal.getAttribute('aria-label')!=='统一命令面板')return;
    event.preventDefault();if(searchOpen){const input=document.getElementById('document-search') as HTMLInputElement|null;input?.focus();input?.select();}else setSearchOpen(true);
  };window.addEventListener('keydown',shortcut);return()=>window.removeEventListener('keydown',shortcut);},[workspace,busy,searchOpen]);
  useEffect(()=>{if(workspace&&selected&&paletteHistory(workspace.uiPreferences.get('commandHistory'))[0]!=='doc:'+selected)workspace.uiPreferences.set('commandHistory',rememberPalette(workspace.uiPreferences.get('commandHistory'),'doc:'+selected));},[workspace,selected]);

  const docs = workspace?.list() ?? []; const meta = docs.find(doc => doc.id === selected); const workspaceName = registry?.entries.get(workspace?.id ?? '')?.name ?? '本地工作区';
  const todayJournal=async()=>{if(!workspace)return;const date=localJournalDate();if(workspace.recycledJournal(date)){setJournalDate(date);return;}setSelected(await workspace.getOrCreateJournal(date));};
  const create = () => run(async () => { if (workspace) setSelected(await workspace.create()); });
  const selectWorkspace = (id: string) => run(async () => { if (!registry || !workspace) return; await workspace.flush(); activate(await registry.select(id)); });
  const remove = () => run(async () => { if (!workspace || !selected) return; setConfirmDelete(false); await workspace.delete(selected); setBench(state=>reconcileWorkbench(state,new Set(workspace.list().map(doc=>doc.id)))); });
  const documentRequired=selected?undefined:'请先打开一篇文档';const pageRequired=documentRequired||(currentTab?.mode==='board'?'请先切换到页面模式':undefined);
  const editorAction=(kind:'find'|'export'|'markdown')=>{if(selected)setEditorCommand({docId:selected,pane:bench.activePane,kind,tick:++commandTick.current});};
  const changeDocumentMode=async(mode:'page'|'board')=>{if(!selected||!workspace)return;const pane=bench.activePane,item=await workspace.openDocument(selected);readDocument(item.store);item.store.history.stopCapturing();setBench(state=>changeTabMode(state,pane,selected,mode));};
  const commands:PaletteCommand[]=workspace?[
    {id:'navigation-history',label:'打开导航历史',keywords:'history navigation 最近访问',disabled:docs.length?undefined:'请先创建文档',run:()=>{window.dispatchEvent(new CustomEvent('open-document-navigation',{detail:workspace.id}));}},
    {id:'navigation-back',label:'导航后退',keywords:'back previous 返回',disabled:(bench[bench.activePane].history?.cursor??-1)>0?undefined:'没有更早的访问记录',run:()=>setBench(state=>navigateHistory(state,state.activePane,(state[state.activePane].history?.cursor??0)-1))},
    {id:'navigation-forward',label:'导航前进',keywords:'forward next',disabled:(bench[bench.activePane].history?.cursor??-1)>=0&&(bench[bench.activePane].history?.cursor??0)<(bench[bench.activePane].history?.entries.length??0)-1?undefined:'没有后续访问记录',run:()=>setBench(state=>navigateHistory(state,state.activePane,(state[state.activePane].history?.cursor??0)+1))},
    {id:'new-document',label:'新建文档',keywords:'create new document 创建',run:async()=>{setSelected(await workspace.create());}},
    {id:'today-journal',label:'打开今日日记',keywords:'journal diary daily 日历',run:todayJournal},
    {id:'journal-calendar',label:'打开日历笔记',keywords:'journal calendar 日期日记',run:()=>setJournalDate(localJournalDate())},
    {id:'templates',label:'从模板新建',keywords:'template 模板文档',run:()=>setTemplateDialog({})},
    {id:'save-template',label:'保存为文档模板',keywords:'save template',disabled:documentRequired,run:()=>{if(selected)setTemplateDialog({saveDocId:selected});}},
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
    {id:'trash',label:'打开回收站',keywords:'trash restore 恢复文档',run:()=>setTrashOpen(true)},
    {id:'organization',label:'管理标签与集合',keywords:'tags collection organize',run:()=>setOrganizationOpen(true)},
  ]:[];
  return <div className="app-shell"><DocumentSidebar onClipper={()=>setClipperOpen(true)} onToday={()=>void run(todayJournal)} onJournal={()=>setJournalDate(localJournalDate())} onAppearance={()=>setThemeOpen(true)} workspace={workspace} workspaces={registry?.list() ?? []} selected={selected} busy={busy} onCreate={() => void create()} onSelect={id => { setSelected(id); setConfirmDelete(false); }} onWorkspace={id => void selectWorkspace(id)} onSearch={() => setSearchOpen(true)} onWorkspaceAction={kind => setDialog({ kind })} onTemplates={()=>setTemplateDialog({})} onTrash={()=>setTrashOpen(true)} onOrganization={() => setOrganizationOpen(true)} onFolder={folder => setDialog({ kind: 'folder', folder })} onImport={()=>setImportDialog({kind:'markdown'})} onNotionImport={()=>setImportDialog({kind:'notion'})} onRestore={() => backupInput.current?.click()} onBackup={() => void run(async () => { if (!workspace) return; const data = await exportBackup(workspace, workspaceName); downloadFile(`${workspaceName.replace(/[\\/:*?"<>|]/g, '_')}.mydoc.json`, JSON.stringify(data), 'application/json'); setNotice('工作区备份已下载'); })} />
    <input ref={markdownInput} type="file" hidden multiple accept=".md,.markdown,.txt,text/markdown,text/plain" aria-label="导入 Markdown 文件" onChange={event => { const files = [...(event.target.files ?? [])]; event.target.value = ''; if(files.length)setImportDialog({kind:'markdown',files}); }} />
    <input ref={backupInput} type="file" hidden accept=".json,application/json" aria-label="恢复工作区备份文件" onChange={event => { const file = event.target.files?.[0]; event.target.value = ''; if (file) void run(async () => { if (!registry || !workspace) return; if (file.size > 50 * 1024 * 1024) throw new Error('备份超过 50 MB，请分工作区备份'); const backup = parseBackup(await file.text()); await workspace.flush(); activate(await registry.restore(backup)); setNotice(`已恢复 ${backup.documents.length} 篇文档和 ${backup.trash?.length??0} 篇回收文档到新工作区`); }); }} />
    <main>{error && <div className="error-banner" role="alert">{error}{!workspace && <button onClick={() => setRetry(value => value + 1)}>重新打开</button>}</div>}{notice && <div className="operation-notice" role="status">{notice}<button aria-label="关闭提示" onClick={() => setNotice('')}>×</button></div>}{workspace && <div className="workspace-status"><SaveStatus persistence={workspace.persistence} />{registry?.persistence.state === 'error' && <SaveStatus persistence={registry.persistence} />}{busy && <span role="status"> · 正在处理…</span>}</div>}
      {!workspace && !error ? <div className="empty-state" role="status">正在打开本地工作区…</div> : workspace ? <>{!!docs.length && <Workbench workspace={workspace} state={bench} onChange={setBench} onCloseTab={(pane,id)=>void run(async()=>{await (await workspace.openDocument(id)).persistence.flush();setBench(state=>closeTab(state,pane,id));})} busy={busy} render={(opened, tab, pane)=>{ const entry=workspace.documents.get(tab.docId); if(!entry)return null;return <DocumentEditor loadWorkspace={id=>registry!.get(id)} key={`${workspace.id}:${pane}:${tab.docId}`} item={opened} meta={entry} workspace={workspace} name={workspaceName} active={bench.activePane===pane} tab={tab} onMode={mode=>setBench(state=>changeTabMode(state,pane,tab.docId,mode))} boardRequest={tab.frameId?{frameId:tab.frameId,tick:tab.tick??0}:undefined} onNavigate={(id, targetWorkspace, frameId, other)=>{ const targetPane:PaneId=other?(pane==='left'?'right':'left'):pane; if(targetWorkspace===workspace.id){if(!workspace.documents.has(id)){setError('引用的文档已删除，正文内容已保留。');return;}setBench(state=>openTab(state,id,targetPane,frameId));setConfirmDelete(false);}else void run(async()=>{if(!registry||!registry.entries.has(targetWorkspace))throw new Error('引用的工作区不存在，正文内容已保留。');const target=await registry.get(targetWorkspace);if(!target.documents.has(id))throw new Error('引用的文档已删除，正文内容已保留。');await workspace.flush();activate(await registry.select(targetWorkspace),id);});}} onMove={folder=>void run(async()=>{workspace.moveDocument(entry.id,folder||null);await workspace.persistence.flush();})} busy={busy} onCommandHandled={()=>setEditorCommand(undefined)} commandRequest={editorCommand?.docId===entry.id&&editorCommand.pane===pane?editorCommand:undefined} onSaveTemplate={()=>setTemplateDialog({saveDocId:entry.id})} onDelete={()=>setConfirmDelete(true)}/>;}} />}{!docs.length && busy && <div className="empty-state" role="status">正在保存变更…</div>}{!docs.length && !busy && <div className="empty-state"><FileText size={32}/><h1>开始你的第一篇文档</h1><p>记录想法与资料，内容自动保存到本机。</p><button className="primary" disabled={busy} onClick={()=>void create()}>新建文档</button></div>}</> : null}
      {confirmDelete && <div className="delete-confirm" role="alert"><span>删除“{meta?.title || '无标题文档'}”？将移入回收站，可随时恢复。</span><button onClick={() => setConfirmDelete(false)}>取消</button><button className="danger" onClick={() => void remove()}>确认删除</button></div>}
    </main>
    {clipperOpen&&workspace&&<WebClipperDialog key={workspace.id} workspace={workspace} onClose={()=>setClipperOpen(false)} onFinish={id=>{setClipperOpen(false);setSelected(id);}} onTrash={()=>{setClipperOpen(false);setTrashOpen(true);}}/>}
    {importDialog&&workspace&&<DocumentImportDialog key={workspace.id} workspace={workspace} initialKind={importDialog.kind} initialFiles={importDialog.files} onClose={()=>setImportDialog(null)} onFinish={id=>{setImportDialog(null);if(id)setSelected(id);}}/>}
    {templateDialog && workspace && <DocumentTemplateDialog key={workspace.id} workspace={workspace} saveDocId={templateDialog.saveDocId} onClose={()=>setTemplateDialog(null)} onCreate={id=>{setTemplateDialog(null);setSelected(id);setNotice('已从模板创建文档');}} />}
    {journalDate&&workspace&&<JournalDialog key={workspace.id} workspace={workspace} initialDate={journalDate} onClose={()=>setJournalDate(null)} onOpen={id=>{setJournalDate(null);setSelected(id);setConfirmDelete(false);}}/>}
    {trashOpen && workspace && <TrashDialog key={workspace.id} workspace={workspace} onClose={()=>setTrashOpen(false)} onRestore={id=>{setTrashOpen(false);setSelected(id);setNotice('文档已恢复，引用将重新生效');}} />}
    {organizationOpen && workspace && <OrganizationManager key={workspace.id} workspace={workspace} onClose={() => setOrganizationOpen(false)} />}
    {themeOpen&&<ThemeDialog initial={appearance.mode} preview={appearance.setMode} save={appearance.save} onClose={()=>setThemeOpen(false)}/>}
    {searchOpen && workspace && <QuickSearch commands={commands} workspace={workspace} onClose={() => setSearchOpen(false)} onSelect={id => { setSelected(id); setConfirmDelete(false); }} />}
    {dialog && registry && workspace && <ResourceDialog key={`${dialog.kind}:${dialog.folder?.id ?? ''}`} title={dialog.kind === 'folder' ? dialog.folder ? '管理文件夹' : '新建文件夹' : dialog.kind === 'create' ? '新建工作区' : dialog.kind === 'rename' ? '重命名工作区' : '删除工作区'} initialName={dialog.kind === 'folder' ? dialog.folder?.name : dialog.kind === 'rename' ? workspaceName : ''} nameLabel={dialog.kind === 'folder' ? '文件夹名称' : '工作区名称'} confirm={dialog.kind === 'delete'} destructive={dialog.kind === 'delete'} onClose={() => setDialog(null)} action={async (name, form) => {
      if (dialog.kind === 'create') { await workspace.flush(); activate(await registry.create(name)); }
      else if (dialog.kind === 'rename') await registry.rename(workspace.id, name);
      else if (dialog.kind === 'delete') {  try { await registry.delete(workspace.id); activate(await registry.select(registry.activeId)); } catch (error) { if (!registry.entries.has(workspace.id)) activate(await registry.select(registry.activeId)); throw error; } }
      else { const parent = String(form.get('parent') || '') || null; if (dialog.folder) { workspace.moveFolder(dialog.folder.id, parent); workspace.renameFolder(dialog.folder.id, name); } else workspace.createFolder(name, parent); await workspace.persistence.flush(); }
    }}>
      {dialog.kind === 'delete' && <p>删除“{workspaceName}”及其中的 {docs.length} 篇文档及 {workspace.trash.size} 篇回收文档？删除后无法恢复，建议先备份。</p>}
      {dialog.kind === 'folder' && <><label className="form-field">上级文件夹<FolderSelect workspace={workspace} name="parent" label="上级文件夹" value={dialog.folder?.parentId} exclude={dialog.folder?.id} /></label>{dialog.folder && <button type="button" className="folder-delete" onClick={() => void run(async () => { workspace.deleteFolder(dialog.folder!.id); await workspace.persistence.flush(); setDialog(null); })}>删除文件夹，保留文档</button>}</>}
    </ResourceDialog>}
  </div>;
}
