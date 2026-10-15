import {markdownExport} from './markdown-export';
import {documentAwareness} from '@mydoc/crdt';
import {documentDragHandle} from './document-drag-handle';
import {documentReading,readingJump} from './document-reading';
import {documentComments,type CommentsSnapshot} from './document-comments';
import {DocumentComments} from './document-comments-ui';
import {documentFootnotes} from './document-footnotes';
import {selectionToolbar} from './selection-toolbar';
import {defaultEditorSettings,type EditorSettings} from './editor-settings';
import { DocumentExportDialog, type ExportRequest } from './document-export-dialog';
import { PdfPreview } from './pdf-preview';
import { createDocumentFind, updateDocumentFind, documentFindKey, emptyDocumentFind } from './document-find';
import { DocumentFindBar } from './document-find-bar';
import{documentExpressions}from'./document-expression';
import { useEffect, useRef, useState } from 'react';
import { EditorState, Selection, TextSelection } from 'prosemirror-state';
import { EditorView, type NodeView } from 'prosemirror-view';
import { baseKeymap, chainCommands, exitCode, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { splitListItem, liftListItem, sinkListItem } from 'prosemirror-schema-list';
import { goToNextCell, tableEditing } from 'prosemirror-tables';
import { ySyncPlugin, ySyncPluginKey, yCursorPlugin } from 'y-prosemirror';
import type { BlockStore } from '@mydoc/editor-framework';
import type { Node as PMNode } from 'prosemirror-model';
import { documentSchema, markdownParser, markdownSerializer, migrateDocument, preserveCodeMetadata } from './document-format';
import { compositionGuard, documentInputRules, markdownEnter } from './document-rules';
import { codeHighlight } from './document-highlight';
import { documentSlashMenu } from './document-slash';
import { codeBlockView } from './document-code';
import { webCardPositionAllowed } from './web-links';
import { documentEmbed, type EmbedSource } from './document-embed';
import { documentWeb } from './document-web';
import { documentMedia, type UploadKind } from './document-media';
import type { WorkspaceAssets } from '@mydoc/nbstore';
import { referenceMenu } from './document-reference';
import type { ReferenceTarget } from './document-links';

function taskItem(node: PMNode, view: EditorView, getPos: () => number | undefined, store: BlockStore): NodeView {
  const dom = document.createElement('li'); const checkbox = document.createElement('input'); const contentDOM = document.createElement('div');
  checkbox.type = 'checkbox'; checkbox.contentEditable = 'false'; checkbox.setAttribute('aria-label', '完成待办');
  dom.append(checkbox, contentDOM);
  const update = (next: PMNode) => {
    if (next.type !== node.type) return false;
    node = next; const task = next.attrs.checked !== null;
    dom.classList.toggle('task-item', task); dom.classList.toggle('task-done', task && next.attrs.checked);
    checkbox.hidden = !task; checkbox.checked = Boolean(next.attrs.checked); return true;
  };
  checkbox.addEventListener('change', () => { const position = getPos(); if (position !== undefined) { store.history.stopCapturing(); view.dispatch(view.state.tr.setNodeMarkup(position, undefined, { ...node.attrs, checked: checkbox.checked })); store.history.stopCapturing(); } });
  update(node);
  return { dom, contentDOM, update, stopEvent: event => event.target === checkbox, ignoreMutation: mutation => !contentDOM.contains(mutation.target) };
}
const SELECTION_HISTORY = Symbol('document-selection');
export function DocumentEditor({ blockRequest, store, documentId, workspaceId = 'local', documents = [], onNavigate = () => {}, referenceRequest = 0, headingRequest, commandRequest, onCommandHandled, assets, embedSource, focusEnabled = true, settings=defaultEditorSettings,onOpenSettings, readOnly=false }: { readOnly?:boolean; blockRequest?:{id:string;tick:number}; settings?:EditorSettings;onOpenSettings?:()=>void; store: BlockStore; focusEnabled?: boolean; assets?: WorkspaceAssets; embedSource?: EmbedSource; documentId: string; workspaceId?: string; documents?: ReferenceTarget[]; onNavigate?: (id: string, workspaceId: string, frameId?:string) => void; onCommandHandled?:()=>void; commandRequest?: { kind: 'find'|'export'|'markdown'; tick:number }; referenceRequest?: number; headingRequest?: { position: number; tick: number } }) {
  const readonlyRef=useRef(readOnly);readonlyRef.current=readOnly;
  const toolbarOptions=useRef({enabled:focusEnabled,settings});  const focusAllowed = useRef(focusEnabled); focusAllowed.current = focusEnabled;
  const host = useRef<HTMLDivElement>(null); const editor = useRef<EditorView | null>(null);
  const [mode, setMode] = useState<'edit' | 'source'>('edit'); const modeRef = useRef(mode); modeRef.current = mode;toolbarOptions.current={enabled:!readOnly&&focusEnabled&&mode==='edit',settings};
  const context = useRef({ source: embedSource, workspaceId, documentId, documents, navigate: onNavigate, notify: (message: string) => setError(message) }); context.current = { source: embedSource, workspaceId, documentId, documents, navigate: onNavigate, notify: message => setError(message) };
  const reading=useRef<ReturnType<typeof documentReading>|null>(null);
  const comments=useRef<ReturnType<typeof documentComments>|null>(null);
  const [commentSnapshot,setCommentSnapshot]=useState<CommentsSnapshot>({rows:[],canAdd:false}),[commentsOpen,setCommentsOpen]=useState(false),[commentDraft,setCommentDraft]=useState(false),[selectedComments,setSelectedComments]=useState<string[]>([]);
  const footnotes=useRef<ReturnType<typeof documentFootnotes>|null>(null);
  const references = useRef<ReturnType<typeof referenceMenu> | null>(null);
  const surfaces = useRef<ReturnType<typeof documentEmbed>|null>(null);
  const embeds = useRef<ReturnType<typeof documentEmbed> | null>(null);
  const web = useRef<ReturnType<typeof documentWeb> | null>(null);
  const media = useRef<ReturnType<typeof documentMedia> | null>(null); const imageInput = useRef<HTMLInputElement>(null); const attachmentInput = useRef<HTMLInputElement>(null); const [uploadProgress, setUploadProgress] = useState('');
  const [exportRequest,setExportRequest]=useState<ExportRequest>();
  const [pdfAssetId,setPdfAssetId]=useState<string>();
  const [find,setFind]=useState(emptyDocumentFind);const findState=useRef(find);findState.current=find;
  const closeFind=(focus=true)=>{const view=editor.current;if(!view)return;updateDocumentFind(view,{open:false});if(focus&&focusAllowed.current&&modeRef.current==='edit')view.focus();};
  const openFind=()=>{const view=editor.current;if(!view||!focusAllowed.current||modeRef.current!=='edit')return;const selection=view.state.selection;const query=view.hasFocus()&&!selection.empty&&selection.$from.sameParent(selection.$to)?view.state.doc.textBetween(selection.from,selection.to).slice(0,256):findState.current.query;updateDocumentFind(view,{open:true,query});requestAnimationFrame(()=>{if(!focusAllowed.current)return;const input=host.current?.closest('.markdown-document')?.querySelector<HTMLInputElement>('.document-find-input input');input?.focus({preventScroll:true});input?.select();});};
  useEffect(()=>{const shortcut=(event:KeyboardEvent)=>{if(event.defaultPrevented||event.isComposing||!focusAllowed.current||modeRef.current!=='edit'||document.querySelector('dialog[open]')||(event.target as HTMLElement)?.closest?.('[role=dialog]'))return;if((event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==='f'){event.preventDefault();openFind();}else if(event.key==='Escape'&&findState.current.open){event.preventDefault();closeFind();}};window.addEventListener('keydown',shortcut);return()=>window.removeEventListener('keydown',shortcut);},[]);
  useEffect(()=>{if(!focusEnabled||mode!=='edit')closeFind(false);},[focusEnabled,mode]);
  const [source, setSource] = useState(''); const [error, setError] = useState('');const [readingError,setReadingError]=useState('');
  useEffect(() => {
    store.history.stopCapturing();
    const fragment = migrateDocument(store); store.history.addTrackedOrigin(ySyncPluginKey);
    let currentSelection: ReturnType<Selection['toJSON']> | null = null;
    const onAdded = ({ stackItem }: { stackItem: { meta: Map<unknown, unknown> } }) => {
      if (focusAllowed.current && !stackItem.meta.has(SELECTION_HISTORY)) stackItem.meta.set(SELECTION_HISTORY, { before: currentSelection, after: currentSelection });
    };
    const onPopped = ({ stackItem, type }: { stackItem: { meta: Map<unknown, unknown> }; type: 'undo' | 'redo' }) => {
      const saved = stackItem.meta.get(SELECTION_HISTORY) as { before: ReturnType<Selection['toJSON']>; after: ReturnType<Selection['toJSON']> } | undefined;
      if (!saved || !focusAllowed.current) return;
      (type === 'undo' ? store.history.redoStack : store.history.undoStack).at(-1)?.meta.set(SELECTION_HISTORY, saved);
      queueMicrotask(() => {
        const view = editor.current; if (!view) return;
        const selection = type === 'undo' ? saved.before : saved.after;
        if (!selection) return;
        try { view.dispatch(view.state.tr.setSelection(Selection.fromJSON(view.state.doc, selection))); }
        catch { view.dispatch(view.state.tr.setSelection(Selection.atEnd(view.state.doc))); }
        if (modeRef.current === 'edit') view.focus();
      });
    };
    store.history.on('stack-item-added', onAdded); store.history.on('stack-item-popped', onPopped);
    const surface=documentEmbed(()=>context.current,store,true);surfaces.current=surface;
    const embed = documentEmbed(() => context.current, store); embeds.current = embed;
    const cards = documentWeb(store, setError); web.current = cards;
    const files = assets ? documentMedia(assets, store, setError, setUploadProgress, setPdfAssetId) : null; media.current = files;
    const reader=documentReading(store,workspaceId,documentId,setReadingError);reading.current=reader;
    const review=documentComments(store,setCommentSnapshot,ids=>{if(editor.current&&(matchMedia('(max-width:1024px)').matches||!!host.current?.closest('.is-split')))readingJump(editor.current);setCommentsOpen(true);setCommentDraft(!ids.length);setSelectedComments(ids);requestAnimationFrame(()=>{if(!ids.length)host.current?.closest('.markdown-document')?.querySelector<HTMLTextAreaElement>('[aria-label="新评论内容"]')?.focus();else host.current?.closest('.markdown-document')?.querySelector<HTMLElement>('[data-thread-id="'+ids[0]+'"]')?.scrollIntoView({block:'nearest'});});},setError);comments.current=review;
    const expressions=documentExpressions(store);const notes=documentFootnotes(store,setError);footnotes.current=notes;
    const reference = referenceMenu(() => context.current, store); references.current = reference;
    const view = new EditorView(host.current!, {
      editable:()=>!readonlyRef.current,
      state: EditorState.create({ schema: documentSchema, plugins: [
        ySyncPlugin(fragment), ...(documentAwareness(store.doc)?[yCursorPlugin(documentAwareness(store.doc)!)]:[]), documentDragHandle(store,()=>toolbarOptions.current.enabled), selectionToolbar(store,()=>toolbarOptions.current,()=>review.open()), createDocumentFind(setFind), compositionGuard,reader.plugin,review.plugin,notes.plugin,expressions.plugin, embed.plugin, surface.plugin, cards.plugin, ...(files ? [files.plugin] : []), reference.plugin, documentSlashMenu(assets ? kind => (kind === 'image' ? imageInput : attachmentInput).current?.click() : undefined, kind => cards.open(kind), embedSource ? () => embed.open() : undefined,embedSource?()=>surface.open():undefined,()=>notes.open()), documentInputRules(), codeHighlight, tableEditing(),
        keymap({
          'Mod-z': () => { store.history.undo(); return true; }, 'Mod-Shift-z': () => { store.history.redo(); return true; }, 'Mod-y': () => { store.history.redo(); return true; },
          'Mod-Alt-m':()=>{review.open();return true;},
          'Mod-b': toggleMark(documentSchema.marks.strong), 'Mod-i': toggleMark(documentSchema.marks.em),
          'Mod-`': toggleMark(documentSchema.marks.code),
          Enter: chainCommands(markdownEnter, (state, dispatch) => {
            const { $from } = state.selection;
            const item = $from.depth > 0 ? $from.node($from.depth - 1) : null;
            return splitListItem(documentSchema.nodes.list_item, { checked: item?.attrs.checked === null ? null : false })(state, dispatch);
          }, (state, dispatch) => state.selection.$from.parent.type === documentSchema.nodes.heading ? baseKeymap.Enter(state, tr => dispatch?.(tr.setBlockType(tr.selection.from, tr.selection.to, documentSchema.nodes.paragraph))) : false, baseKeymap.Enter),
          'Shift-Enter': (state, dispatch) => { dispatch?.(state.tr.replaceSelectionWith(documentSchema.nodes.hard_break.create()).scrollIntoView()); return true; },
          'Mod-Enter': exitCode,
          Tab: chainCommands(goToNextCell(1), sinkListItem(documentSchema.nodes.list_item), (state, dispatch) => { if (state.selection.$from.parent.type !== documentSchema.nodes.code_block) return false; dispatch?.(state.tr.insertText('  ')); return true; }),
          'Shift-Tab': chainCommands(goToNextCell(-1), liftListItem(documentSchema.nodes.list_item)),
        }), keymap(baseKeymap),
      ] }),
      dispatchTransaction(this: EditorView, tr) {
        if(readonlyRef.current&&tr.docChanged&&!tr.getMeta(ySyncPluginKey))return;
        if(tr.scrolledIntoView&&!tr.getMeta(ySyncPluginKey))readingJump(this);
        const local = tr.docChanged && !tr.getMeta(ySyncPluginKey);
        const structural = local && (!this.state.selection.empty || tr.steps.some(step => {
          const json = step.toJSON();
          if (json.stepType !== 'replace') return true;
          if (json.slice?.content?.some((node: { type: string }) => node.type !== 'text')) return true;
          return typeof json.from === 'number' && typeof json.to === 'number' && this.state.doc.resolve(json.from).parent !== this.state.doc.resolve(json.to).parent;
        }));
        if (structural) store.history.stopCapturing();
        currentSelection = this.state.selection.toJSON();
        this.updateState(this.state.apply(tr));
        if (tr.docChanged && !tr.getMeta(ySyncPluginKey) && !store.history.undoing && !store.history.redoing) {
          const entry = store.history.undoStack.at(-1)?.meta.get(SELECTION_HISTORY) as { after: ReturnType<Selection['toJSON']> } | undefined;
          if (entry) entry.after = this.state.selection.toJSON();
        }
        if (structural) store.history.stopCapturing();
        currentSelection = this.state.selection.toJSON();
        if (tr.docChanged && modeRef.current === 'source') setSource(markdownSerializer.serialize(this.state.doc));
      },
      handleDOMEvents: { beforeinput(view, event) {
        const input = event as InputEvent;
        if (input.isComposing || view.state.selection.$from.parent.type !== documentSchema.nodes.code_block || !input.data?.includes('\n')) return false;
        // Preserve multiline native insertions inside the single code contentDOM.
        event.preventDefault(); view.dispatch(view.state.tr.insertText(input.data).scrollIntoView()); return true;
      } },
      handlePaste(this: EditorView, view, event) {
        if (event.clipboardData?.getData('text/html')?.match(/data-web-card|data-doc-embed/) && !webCardPositionAllowed(view.state.selection)) { setError('请在代码和表格外粘贴网页卡片。'); return true; }
        const text = event.clipboardData?.getData('text/plain');
        if (!text || view.state.selection.$from.parent.type.spec.code) return false;
        // 文本剪贴板按 Markdown 解析；富文本 HTML 交给 ProseMirror 正常处理。
        if (event.clipboardData?.getData('text/html')) return false;
        try { const parsed = markdownParser.parse(text); view.dispatch(view.state.tr.replaceSelection(parsed.slice(0, parsed.content.size)).scrollIntoView()); setError(''); return true; }
        catch { setError('Markdown 粘贴解析失败，已保留原文档。'); return false; }
      },
      nodeViews: {...notes.nodeViews,...expressions.nodeViews, surface_ref:(node,view,getPos)=>surface.nodeView(node,view,getPos), doc_embed: (node, view, getPos) => embed.nodeView(node, view, getPos), web_card: (node, view, getPos) => cards.nodeView(node, view, getPos), ...(files ? { image: (node: PMNode) => files.nodeView(node), attachment: (node: PMNode) => files.nodeView(node) } : {}), doc_reference: node => reference.nodeView(node), code_block: (node, view, getPos) => codeBlockView(node, view, getPos, store, documentId), list_item: (node, view, getPos) => taskItem(node, view, getPos, store) },
      attributes: { role: 'textbox', 'aria-label': '文档正文', 'aria-multiline': 'true', 'data-placeholder': '输入文字，/ 选择格式，@ 引用文档…' },
    });
    editor.current = view; currentSelection = view.state.selection.toJSON();
    const fileHost=view.dom.closest('.markdown-document');
    const insertFiles=(event:Event)=>{if(readonlyRef.current)return;if(event.target!==fileHost)return;event.preventDefault();if(modeRef.current!=='edit'){setError('请切回编辑模式，再添加文件。');return;}void files?.upload((event as CustomEvent<File[]>).detail,'auto');};
    fileHost?.addEventListener('mydoc:insert-files',insertFiles);

    if (focusAllowed.current && !view.state.doc.textContent) { view.focus(); }
    const navigateBlock = () => { const id = new URLSearchParams(location.hash.slice(1)).get('block'); if(id)readingJump(view); if (id) queueMicrotask(() => view.dom.querySelector(`#${CSS.escape(`code-${id}`)}`)?.scrollIntoView({ block: 'center' })); };
    navigateBlock(); window.addEventListener('hashchange', navigateBlock);
    return () => { fileHost?.removeEventListener('mydoc:insert-files',insertFiles);window.removeEventListener('hashchange', navigateBlock); store.history.stopCapturing(); store.history.off('stack-item-added', onAdded); store.history.off('stack-item-popped', onPopped); editor.current = null; references.current = null; media.current = null; web.current = null; embeds.current = null;surfaces.current=null;footnotes.current=null;comments.current=null;reading.current=null; view.destroy(); };
  }, [store, documentId, workspaceId, assets, embedSource]);
  useEffect(()=>{editor.current?.setProps({editable:()=>!readonlyRef.current});},[readOnly]);
  useEffect(()=>{const view=editor.current;if(view)view.dispatch(view.state.tr.setMeta('toolbar-options',true));},[focusEnabled,mode,settings.width,settings.fontSize,settings.lineHeight,settings.selectionToolbar,settings.mobileToolbar]);
  useEffect(()=>{if(!blockRequest||!focusEnabled)return;setMode('edit');const frame=requestAnimationFrame(()=>{const view=editor.current;if(!view)return;const block=view.dom.querySelector(`#${CSS.escape('code-'+blockRequest.id)}`);if(!block){setError('链接对应的正文块已不存在');return;}setError('');readingJump(view);block.scrollIntoView({block:'center'});});return()=>cancelAnimationFrame(frame);},[blockRequest?.id,blockRequest?.tick,focusEnabled]);
  useEffect(() => { references.current?.refresh(); embeds.current?.refresh();surfaces.current?.refresh(); }, [documents]);
  useEffect(() => { if (!referenceRequest) return; setMode('edit'); requestAnimationFrame(() => references.current?.open()); }, [referenceRequest]);
  useEffect(() => { if (!headingRequest) return; setMode('edit'); requestAnimationFrame(() => { const view = editor.current; const position = headingRequest.position;if(view)readingJump(view); if (!view || view.state.doc.nodeAt(position)?.type !== documentSchema.nodes.heading) return; view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(position + 1)))); (view.nodeDOM(position) as HTMLElement | null)?.scrollIntoView({ block: 'center', behavior: 'smooth' }); view.focus(); }); }, [headingRequest]);
  const switchMode = (next: 'edit' | 'source') => {
    if (next === 'source' && editor.current) setSource(markdownSerializer.serialize(editor.current.state.doc));
    setMode(next); if (next === 'edit') queueMicrotask(() => editor.current?.focus());
  };
  const updateSource = (value: string) => {
    if(readonlyRef.current)return;
    const view = editor.current; if (!view) return;
    try { const parsed = preserveCodeMetadata(view.state.doc, markdownParser.parse(value)); view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, parsed.content)); setSource(value); setError(''); }
    catch { setSource(value); setError('Markdown 解析失败，请检查源码。'); }
  };
  const [exportingMarkdown,setExportingMarkdown]=useState(false);
  const markdownExportBusy=useRef(false);
  const download = async () => {
    if (!editor.current||markdownExportBusy.current) return;
    markdownExportBusy.current=true;setExportingMarkdown(true);setError('');
    try{
      const doc=mode==='source'?markdownParser.parse(source):editor.current.state.doc;
      const result=await markdownExport(doc,store.meta.get('title')||'无标题文档',assets,mode==='source'?source:undefined);
      const url=URL.createObjectURL(new Blob([result.bytes as BlobPart],{type:result.type}));
      const link=document.createElement('a');link.href=url;link.download=result.name;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
    }catch(error){setError(error instanceof Error?error.message:'Markdown 导出失败');}
    finally{markdownExportBusy.current=false;setExportingMarkdown(false);}
  };
  useEffect(()=>{if(!commandRequest||!focusAllowed.current)return;onCommandHandled?.();
    if(commandRequest.kind==='markdown'){download();return;}
    if(modeRef.current==='source'&&error){setError('请先修正 Markdown 源码，再执行查找或导出。');return;}
    if(commandRequest.kind==='export'){const view=editor.current;if(view)setExportRequest({doc:view.state.doc,title:store.meta.get('title')||'无标题文档'});}
    else {setMode('edit');requestAnimationFrame(openFind);}
  },[commandRequest?.tick]);
  const selectedFiles = (input: HTMLInputElement, kind: UploadKind) => { if(readonlyRef.current)return; const files = [...(input.files ?? [])]; input.value = ''; if (files.length) { setMode('edit'); void media.current?.upload(files, kind); } };
  return <div onClickCapture={event=>{if(readOnly&&(event.target as HTMLElement).closest('button,input[type=checkbox]')&&!(event.target as HTMLElement).closest('.document-find,.document-tools,.document-options,dialog,.pdf-preview,.document-comments-panel,.comment-controls,[data-readonly-action]')){event.preventDefault();event.stopPropagation();}}} onKeyDownCapture={event=>{if(readOnly&&(event.metaKey||event.ctrlKey)&&!['c','a','f'].includes(event.key.toLowerCase())){event.preventDefault();event.stopPropagation();}}} className={"markdown-document"+(commentsOpen&&mode==='edit'?' comments-open':'')}>
    {pdfAssetId&&assets&&<PdfPreview assetId={pdfAssetId} assets={assets} onClose={()=>{setPdfAssetId(undefined);requestAnimationFrame(()=>{if(focusAllowed.current)editor.current?.focus();});}}/>}
    {exportRequest&&<DocumentExportDialog request={exportRequest} assets={assets} resolveTitle={(id,ws)=>ws===workspaceId?documents.find(d=>d.id===id)?.title:undefined} onClose={()=>{setExportRequest(undefined);requestAnimationFrame(()=>{if(!focusAllowed.current)return;if(modeRef.current==='edit')editor.current?.focus();else host.current?.closest('.markdown-document')?.querySelector<HTMLTextAreaElement>('.markdown-source')?.focus();});}} />}
    <details className="document-options"><summary aria-label="文档选项" title="文档选项">···</summary><div className="document-tools"><div role="group" aria-label="编辑模式"><button aria-pressed={mode === 'edit'} onClick={() => switchMode('edit')}>编辑</button><button aria-pressed={mode === 'source'} onClick={() => switchMode('source')}>Markdown 源码</button></div><button disabled={mode!=='edit'} onClick={event=>{(event.currentTarget.closest('details') as HTMLDetailsElement).open=false;openFind();}}>页内查找</button><button disabled={mode==='source'&&!!error} onClick={event=>{(event.currentTarget.closest('details') as HTMLDetailsElement).open=false;const view=editor.current;if(view)setExportRequest({doc:view.state.doc,title:store.meta.get('title')||'无标题文档'});}}>导出 PDF / HTML</button><button disabled={exportingMarkdown} onClick={()=>void download()}>{exportingMarkdown?'正在导出…':'导出 Markdown'}</button>{readingError&&<button onClick={()=>reading.current?.retry()}>重试保存阅读位置</button>}<button disabled={readOnly||mode!=='edit'} onClick={event=>{(event.currentTarget.closest('details') as HTMLDetailsElement).open=false;footnotes.current?.open();}}>插入脚注</button>{onOpenSettings&&<button onClick={event=>{(event.currentTarget.closest('details') as HTMLDetailsElement).open=false;onOpenSettings();}}>编辑器设置</button>}</div></details>
    {assets && !readOnly && <div className="media-upload-controls" role="group" aria-label="本地文件"><button disabled={readOnly||!!uploadProgress} onClick={() => imageInput.current?.click()}>上传图片</button><button disabled={readOnly||!!uploadProgress} onClick={() => attachmentInput.current?.click()}>添加附件</button><button onClick={() => { setMode('edit'); web.current?.open('bookmark'); }}>添加书签</button><button onClick={() => { setMode('edit'); web.current?.open('embed'); }}>嵌入网页</button><button onClick={() => { setMode('edit'); embeds.current?.open(); }}>嵌入文档</button><button onClick={()=>{setMode('edit');surfaces.current?.open();}}>引用画布</button><span role="status">{uploadProgress}</span><input ref={imageInput} type="file" hidden multiple accept="image/png,image/jpeg,image/gif,image/webp" aria-label="上传本地图片" onChange={event => selectedFiles(event.target, 'image')} /><input ref={attachmentInput} type="file" hidden multiple aria-label="添加本地附件" onChange={event => selectedFiles(event.target, 'attachment')} /></div>}
    {find.open && focusEnabled && mode==='edit' && <DocumentFindBar state={find} view={()=>editor.current} onClose={()=>closeFind()} />}
    {error && <div className="error-banner" role="alert">{error}</div>}{readingError&&<div className="error-banner" role="alert">{readingError}</div>}
    <div className="comment-controls"><button type="button" aria-expanded={commentsOpen&&mode==='edit'} disabled={mode!=='edit'} onClick={()=>{setCommentsOpen(!commentsOpen);if(!commentsOpen&&editor.current&&(matchMedia('(max-width:1024px)').matches||!!host.current?.closest('.is-split')))readingJump(editor.current);if(!commentsOpen&&(matchMedia('(max-width:1024px)').matches||!!host.current?.closest('.is-split')))requestAnimationFrame(()=>host.current?.closest('.markdown-document')?.querySelector('.document-comments-panel,.comment-controls,[data-readonly-action]')?.scrollIntoView({block:'start'}));if(commentsOpen){comments.current?.cancel();setCommentDraft(false);}}}>评论 ({commentSnapshot.rows.length})</button><button type="button" disabled={readOnly||mode!=='edit'||!commentSnapshot.canAdd} onClick={()=>comments.current?.open()}>添加正文评论</button></div>
    <div className="document-review-layout"><div ref={host} className="document-editor" hidden={mode !== 'edit'} />
    {mode==='edit'&&<DocumentComments readOnly={readOnly} snapshot={commentSnapshot} controller={comments.current} open={commentsOpen} draft={commentDraft} selected={selectedComments} onClose={()=>{setCommentsOpen(false);setCommentDraft(false);comments.current?.cancel();if(focusAllowed.current)editor.current?.focus();}} onDraftDone={()=>setCommentDraft(false)} onFilterChange={()=>setSelectedComments([])}/>}</div>
    {mode === 'source' && <textarea readOnly={readOnly} className="markdown-source" aria-label="Markdown 源码" value={source} onChange={event => updateSource(event.target.value)} spellCheck={false} />}
  </div>;
}
