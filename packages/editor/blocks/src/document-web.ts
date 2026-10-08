import { Plugin, TextSelection, type SelectionBookmark } from 'prosemirror-state';
import type { EditorView, NodeView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import type { BlockStore } from '@mydoc/editor-framework';
import { documentSchema } from './document-format';
import { readWebMetadata, validWebCard, webUrl, webCardPositionAllowed, type WebCardAttrs, type WebCardKind } from './web-links';
export function documentWeb(store: BlockStore, notify: (text: string) => void) {
  let editor: EditorView; let dialog: HTMLDialogElement | null = null; let bookmark: SelectionBookmark | null = null; let editingPosition: number | undefined; let abort: AbortController | null = null; let request = 0;
  const close = (focus = true) => { request++; abort?.abort(); abort = null; dialog?.remove(); dialog = null; bookmark = null; editingPosition = undefined; if (focus) editor.focus(); };
  const button = (label: string, action: () => void) => { const element = document.createElement('button'); element.type = 'button'; element.textContent = label; element.addEventListener('click', action); return element; };
  const mutate = (position: number, action: 'delete' | WebCardKind) => { const node = editor.state.doc.nodeAt(position); if (node?.type !== documentSchema.nodes.web_card) return; store.history.stopCapturing(); const tr = action === 'delete' ? editor.state.tr.delete(position, position + node.nodeSize) : editor.state.tr.setNodeMarkup(position, undefined, { ...node.attrs, kind: action }); editor.dispatch(tr.scrollIntoView()); store.history.stopCapturing(); editor.focus(); };
  function open(kind: WebCardKind, position?: number) {
    if (editor.composing || !webCardPositionAllowed(editor.state.selection)) { notify('请在正文中选择网页卡片插入位置（代码和表格外）'); return; }
    close(false); bookmark = editor.state.selection.getBookmark();
    const original = position !== undefined ? editor.state.doc.nodeAt(position) : null;
    const existing = original?.type === documentSchema.nodes.web_card ? original : null; const existingId = existing?.attrs.id; editingPosition = existing ? position : undefined;
    dialog = document.createElement('dialog'); dialog.className = 'web-card-dialog'; dialog.setAttribute('aria-labelledby', 'web-card-dialog-title'); const form = document.createElement('form');
    const heading = document.createElement('h2'); heading.id = 'web-card-dialog-title'; heading.textContent = existing ? '编辑网页卡片' : kind === 'bookmark' ? '添加书签' : '嵌入网页'; form.append(heading);
    function field(label: string, value: string, multiline = false) { const wrapper = document.createElement('label'); wrapper.className = 'form-field'; wrapper.append(document.createTextNode(label)); const input = multiline ? document.createElement('textarea') : document.createElement('input'); input.value = value; input.setAttribute('aria-label', label); wrapper.append(input); form.append(wrapper); return input; }
    const url = field('网页地址', existing?.attrs.url || '') as HTMLInputElement; url.type = 'url'; url.required = true; url.placeholder = 'https://…'; url.maxLength = 8192;
    const title = field('卡片标题', existing?.attrs.title || ''); title.maxLength = 500; const description = field('网页摘要', existing?.attrs.description || '', true); description.maxLength = 2000;
    const hint = document.createElement('p'); hint.className = 'web-card-hint'; hint.textContent = '可读取网页标题与摘要；网站限制跨域读取时，可手动填写。'; form.append(hint);
    const status = document.createElement('p'); status.className = 'web-card-hint'; status.role = 'status'; form.append(status);
    const actions = document.createElement('div'); actions.className = 'web-card-dialog-actions'; const submit = document.createElement('button'); submit.type = 'submit'; submit.className = 'primary'; submit.textContent = existing ? '保存卡片' : '插入卡片';
    const metadata = button('读取网页信息', () => { void fetchMetadata(); }); actions.append(metadata, button('取消', () => close()), submit); form.append(actions); dialog.append(form); document.body.append(dialog);
    url.addEventListener('input', () => { request++; abort?.abort(); metadata.disabled = submit.disabled = false; status.textContent = ''; });
    async function fetchMetadata() {
      let target: string; try { target = webUrl(url.value); } catch { status.role = 'alert'; status.textContent = '请输入有效的 HTTP 或 HTTPS 网址'; return; }
      abort?.abort(); const controller = new AbortController(); abort = controller; const generation = ++request; const timer = setTimeout(() => controller.abort(), 8000); metadata.disabled = submit.disabled = true; status.role = 'status'; status.textContent = '正在读取网页信息…';
      try { const data = await readWebMetadata(target, controller.signal); if (generation !== request || !dialog) return; title.value = data.title; description.value = data.description; status.textContent = data.title || data.description ? '网页信息已读取，确认后保存。' : '网页未提供标题或摘要，可手动填写。'; }
      catch { if (generation === request && dialog) { status.role = 'alert'; status.textContent = '无法读取网页信息，可能是网络或跨域限制。可重试或手动填写后插入。'; } }
      finally { clearTimeout(timer); if (generation === request && dialog) metadata.disabled = submit.disabled = false; }
    }
    form.addEventListener('submit', event => {
      event.preventDefault(); let target: string; try { target = webUrl(url.value); } catch { status.role = 'alert'; status.textContent = '请输入有效的 HTTP 或 HTTPS 网址'; return; }
      const attrs: WebCardAttrs = { id: existingId || crypto.randomUUID(), kind, url: target, title: title.value.trim() || new URL(target).hostname, description: description.value.trim() };
      if (!validWebCard(attrs)) { status.role = 'alert'; status.textContent = '卡片信息过长，请缩短标题或摘要'; return; }
      let tr = editor.state.tr;
      if (existing) { const found = editingPosition; if (found === undefined || editor.state.doc.nodeAt(found)?.type !== documentSchema.nodes.web_card) { status.role = 'alert'; status.textContent = '原卡片已删除，请取消后重新插入'; return; } tr = tr.setNodeMarkup(found, undefined, attrs); }
      else { if (!bookmark) return; tr = tr.setSelection(bookmark.resolve(editor.state.doc)).replaceSelectionWith(documentSchema.nodes.web_card.create(attrs)); const end = tr.selection.to; if (tr.doc.resolve(end).parent === tr.doc) { tr.insert(end, documentSchema.nodes.paragraph.create()); tr.setSelection(TextSelection.near(tr.doc.resolve(end + 1))); } }
      bookmark = null; store.history.stopCapturing(); editor.dispatch(tr.scrollIntoView()); store.history.stopCapturing(); close();
    });
    dialog.addEventListener('cancel', event => { event.preventDefault(); close(); }); dialog.showModal(); url.focus();
  }
  function nodeView(node: PMNode, _view: EditorView, getPos: () => number | undefined): NodeView {
    const dom = document.createElement('div'); dom.className = 'web-card'; dom.contentEditable = 'false'; dom.setAttribute('role', 'group'); dom.setAttribute('aria-label', node.attrs.kind === 'embed' ? '嵌入网页卡片' : '书签卡片');
    const attrs = node.attrs as WebCardAttrs; const valid = validWebCard(attrs); const info = document.createElement('div'); info.className = 'web-card-info'; const heading = document.createElement('strong'); heading.textContent = attrs.title || attrs.url; info.append(heading);
    if (attrs.description) { const description = document.createElement('p'); description.textContent = attrs.description; info.append(description); }
    const address = document.createElement('span'); address.className = 'web-card-address'; address.textContent = valid ? new URL(attrs.url).hostname : '网页地址无效'; info.append(address); dom.append(info);
    const actions = document.createElement('div'); actions.className = 'web-card-actions';
    if (valid) { const link = document.createElement('a'); link.href = webUrl(attrs.url); link.target = '_blank'; link.rel = 'noopener noreferrer'; link.textContent = '打开网页'; actions.append(link); }
    actions.append(button('编辑卡片', () => { const position = getPos(); if (position !== undefined) open(attrs.kind, position); }), button(attrs.kind === 'embed' ? '转为书签' : '转为嵌入', () => { const position = getPos(); if (position !== undefined) mutate(position, attrs.kind === 'embed' ? 'bookmark' : 'embed'); }), button('删除卡片', () => { const position = getPos(); if (position !== undefined) mutate(position, 'delete'); })); dom.append(actions);
    let timer: ReturnType<typeof setTimeout> | undefined; let removed = false;
    if (attrs.kind === 'embed' && valid) {
      const body = document.createElement('div'); body.className = 'web-embed-body'; const status = document.createElement('p'); status.className = 'web-card-hint'; status.role = 'status'; status.textContent = '加载后连接此网站。部分网站限制嵌入，可使用“打开网页”。';
      let frame: HTMLIFrameElement | null = null;
      const load = button('加载网页', () => {
        if (frame) { frame.remove(); frame = null; clearTimeout(timer); load.textContent = '加载网页'; status.textContent = '已停止嵌入网页。'; editor.focus(); return; }
        frame = document.createElement('iframe'); frame.src = webUrl(attrs.url); frame.title = attrs.title || '嵌入网页'; frame.referrerPolicy = 'no-referrer'; frame.setAttribute('sandbox', 'allow-scripts allow-forms allow-popups'); frame.loading = 'lazy'; body.prepend(frame); load.textContent = '停止加载'; status.textContent = '正在加载网页…';
        const activeFrame = frame; frame.addEventListener('load', () => { if (removed || frame !== activeFrame) return; clearTimeout(timer); status.textContent = '已尝试加载网页。若无法显示，请使用“打开网页”；网站可能限制嵌入。'; });
        editor.focus();
        timer = setTimeout(() => { if (!removed && frame) status.textContent = '网页加载时间较长，可停止后重试，或使用“打开网页”。'; }, 10000);
      }); body.append(load, status); dom.append(body);
    }
    return { dom, stopEvent: event => !!(event.target as HTMLElement)?.closest('button, a, iframe'), ignoreMutation: () => true, destroy() { removed = true; clearTimeout(timer); } };
  }
  const plugin = new Plugin({ state: { init: () => null, apply(tr) { if (bookmark) bookmark = bookmark.map(tr.mapping); if (editingPosition !== undefined) { const result = tr.mapping.mapResult(editingPosition); editingPosition = result.deleted ? undefined : result.pos; } return null; } }, view(view) { editor = view; return { update(view) { editor = view; }, destroy() { close(false); } }; } });
  return { plugin, nodeView, open };
}
