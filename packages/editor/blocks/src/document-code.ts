import type { Node as PMNode } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import type { EditorView, NodeView } from 'prosemirror-view';
import type { BlockStore } from '@mydoc/editor-framework';

export const codeLanguages = [
  ['', '纯文本'], ['javascript', 'JavaScript'], ['typescript', 'TypeScript'], ['sql', 'SQL'],
  ['markup', 'HTML'], ['css', 'CSS'], ['json', 'JSON'], ['python', 'Python'], ['bash', 'Shell'],
  ['java', 'Java'], ['c', 'C'], ['cpp', 'C++'], ['csharp', 'C#'], ['go', 'Go'], ['rust', 'Rust'],
  ['yaml', 'YAML'], ['markdown', 'Markdown'],
];
const aliases: Record<string, string> = { js: 'javascript', ts: 'typescript', html: 'markup', sh: 'bash', py: 'python', yml: 'yaml' };
const icons: Record<string, string> = {
  copy: '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/>',
  collapse: '<rect x="4" y="3" width="10" height="7" rx="2"/><path d="M4 15h16M4 20h10"/>',
  comment: '<path d="M21 11a9 9 0 0 1-9 9 10 10 0 0 1-4-1l-5 2 2-5a10 10 0 0 1-1-4 9 9 0 1 1 17-1Z"/>',
  more: '<circle cx="12" cy="5" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="12" cy="19" r="1"/>',
  wrap: '<path d="M3 5h17M3 10h13a4 4 0 0 1 0 8H9m3-3-3 3 3 3"/>',
  numbers: '<path d="M9 5h12M9 12h12M9 19h12M3 3v4M2 10h3l-3 4h3M2 17h3v4H2"/>',
  link: '<path d="m10 13 4-4M9 16l-2 2a4 4 0 0 1-6-6l4-4a4 4 0 0 1 6 0m2 0 2-2a4 4 0 0 1 6 6l-4 4a4 4 0 0 1-6 0"/>',
  duplicate: '<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M15 9V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h4m4 0h5m-2-2 2 2-2 2"/>',
  delete: '<path d="M3 6h18M9 6V3h6v3M5 6l1 15h12l1-15M10 10v7M14 10v7"/>',
};
function icon(name: string) { const span = document.createElement('span'); span.className = 'code-icon'; span.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name]}</svg>`; return span; }
type Comment = { id: string; text: string; createdAt: number };

export function codeBlockView(node: PMNode, view: EditorView, getPos: () => number | undefined, store: BlockStore, documentId: string): NodeView {
  const dom = document.createElement('div'); dom.className = 'code-block'; let localId: string = node.attrs.id || crypto.randomUUID();
  const toolbar = document.createElement('div'); toolbar.className = 'code-toolbar'; toolbar.contentEditable = 'false'; toolbar.setAttribute('role', 'group'); toolbar.setAttribute('aria-label', '代码块工具栏');
  const language = document.createElement('button'); language.type = 'button'; language.className = 'code-language'; language.setAttribute('aria-label', '选择代码语言'); language.setAttribute('aria-expanded', 'false');
  const actions = document.createElement('div'); actions.className = 'code-actions'; toolbar.append(language, actions);
  const pre = document.createElement('pre'); const contentDOM = document.createElement('code'); const gutter = document.createElement('span');
  gutter.className = 'code-line-numbers'; gutter.contentEditable = 'false'; gutter.setAttribute('aria-hidden', 'true'); pre.append(gutter, contentDOM);
  const status = document.createElement('div'); status.className = 'code-status'; status.contentEditable = 'false'; status.setAttribute('role', 'status'); status.hidden = true;
  const comments = document.createElement('div'); comments.className = 'code-comments'; comments.contentEditable = 'false'; comments.hidden = true;
  dom.append(toolbar, pre, status, comments);
  let popup: HTMLDivElement | null = null; let popupTrigger: HTMLButtonElement | null = null; let statusTimer: ReturnType<typeof setTimeout>; let frame = 0;
  let commentsOpen = false;
  const patch = (attrs: Record<string, unknown>) => { const pos = getPos(); if (pos === undefined) return; store.history.stopCapturing(); view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, id: localId, ...attrs })); store.history.stopCapturing(); };
  const notify = (message: string) => { clearTimeout(statusTimer); status.textContent = message; status.hidden = false; statusTimer = setTimeout(() => { status.hidden = true; }, 3000); };
  const copyText = async (text: string) => { try { await navigator.clipboard.writeText(text); notify('已复制'); } catch { notify('复制失败，请允许剪贴板访问后重试'); } };
  const closePopup = (focus = false) => { const trigger = popupTrigger; popup?.remove(); popup = null; popupTrigger = null; trigger?.setAttribute('aria-expanded', 'false'); if (focus) trigger?.focus(); };
  const positionPopup = () => { if (!popup || !popupTrigger) return; const bounds = popupTrigger.getBoundingClientRect(); const width = Math.min(280, innerWidth - 24); popup.style.width = `${width}px`; popup.style.left = `${Math.max(12, Math.min(bounds.left, innerWidth - width - 12))}px`; const height = Math.min(popup.scrollHeight, 330, innerHeight - 24); popup.style.maxHeight = `${height}px`; popup.style.top = `${bounds.bottom + height + 6 < innerHeight ? bounds.bottom + 6 : Math.max(12, bounds.top - height - 6)}px`; };
  const openPopup = (trigger: HTMLButtonElement, label: string) => { if (popupTrigger === trigger) { closePopup(true); return null; } closePopup(); popupTrigger = trigger; trigger.setAttribute('aria-expanded', 'true'); popup = document.createElement('div'); popup.className = 'code-popup'; popup.setAttribute('role', 'dialog'); popup.setAttribute('aria-label', label); document.body.append(popup);
    popup.addEventListener('keydown', event => { if (event.key === 'Escape') { event.preventDefault(); closePopup(true); } if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { const buttons = [...popup!.querySelectorAll<HTMLButtonElement>('button')]; const current = buttons.indexOf(document.activeElement as HTMLButtonElement); buttons[(current + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length]?.focus(); event.preventDefault(); } });
    return popup;
  };
  const button = (label: string, name: string, action: () => void) => { const result = document.createElement('button'); result.type = 'button'; result.title = label; result.setAttribute('aria-label', label); result.append(icon(name)); result.addEventListener('mousedown', event => event.preventDefault()); result.addEventListener('click', action); return result; };
  const collapse = button('折叠代码', 'collapse', () => { patch({ collapsed: !node.attrs.collapsed }); });
  const commentButton = button('代码块评论', 'comment', () => { commentsOpen = !commentsOpen; renderComments(); if (commentsOpen) comments.querySelector('textarea')?.focus(); });
  const more = button('代码块更多选项', 'more', () => {
    const panel = openPopup(more, '代码块更多选项'); if (!panel) return;
    const row = (label: string, name: string, action: () => void, checked?: boolean) => {
      const item = button(label, name, action); item.className = 'code-menu-row'; const text = document.createElement('span'); text.textContent = label; item.append(text);
      if (checked !== undefined) { item.setAttribute('role', 'switch'); item.setAttribute('aria-checked', String(checked)); const toggle = document.createElement('span'); toggle.className = 'code-switch'; item.append(toggle); }
      panel.append(item); return item;
    };
    const wrap = row('自动换行', 'wrap', () => { patch({ wrap: !node.attrs.wrap }); wrap.setAttribute('aria-checked', String(node.attrs.wrap)); }, node.attrs.wrap);
    const numbers = row('显示行号', 'numbers', () => { patch({ lineNumbers: !node.attrs.lineNumbers }); numbers.setAttribute('aria-checked', String(node.attrs.lineNumbers)); }, node.attrs.lineNumbers);
    const separator = () => { const line = document.createElement('hr'); panel.append(line); }; separator();
    row('复制代码块链接', 'link', () => { patch({}); const url = new URL(location.href); const params = new URLSearchParams(location.hash.slice(1)); params.set('doc', documentId); params.set('block', localId); url.hash = params.toString(); void copyText(url.href); closePopup(); });
    row('复制代码块', 'duplicate', () => { const pos = getPos(); if (pos === undefined) return; closePopup(); store.history.stopCapturing(); const duplicate = node.type.create({ ...node.attrs, id: crypto.randomUUID(), comments: [] }, node.content); const tr = view.state.tr.insert(pos + node.nodeSize, duplicate); tr.setSelection(TextSelection.create(tr.doc, pos + node.nodeSize + 1)); view.dispatch(tr.scrollIntoView()); store.history.stopCapturing(); view.focus(); });
    separator(); row('删除代码块', 'delete', () => { const pos = getPos(); if (pos === undefined) return; closePopup(); store.history.stopCapturing(); view.dispatch(view.state.tr.delete(pos, pos + node.nodeSize).scrollIntoView()); store.history.stopCapturing(); view.focus(); });
    positionPopup(); panel.querySelector<HTMLButtonElement>('button')?.focus();
  });
  actions.append(button('复制代码', 'copy', () => { void copyText(node.textContent); }), collapse, commentButton, more);
  language.addEventListener('mousedown', event => event.preventDefault()); language.addEventListener('click', () => {
    const panel = openPopup(language, '代码语言'); if (!panel) return;
    const search = document.createElement('input'); search.type = 'search'; search.placeholder = '搜索语言'; search.setAttribute('aria-label', '搜索代码语言');
    const list = document.createElement('div'); list.className = 'code-language-list'; const render = () => { list.replaceChildren(); const filtered = codeLanguages.filter(([id, label]) => `${id} ${label} ${Object.entries(aliases).filter(([, value]) => value === id).map(([alias]) => alias).join(' ')} ${id ? '' : 'plain text'}`.toLowerCase().includes(search.value.toLowerCase()));
      for (const [id, label] of filtered) { const option = document.createElement('button'); option.type = 'button'; option.textContent = label; option.setAttribute('aria-pressed', String(id === (aliases[node.attrs.params] || node.attrs.params))); option.addEventListener('click', () => { patch({ params: id }); closePopup(); view.focus(); }); list.append(option); }
      if (!filtered.length) { const empty = document.createElement('p'); empty.textContent = '没有匹配的语言'; list.append(empty); } positionPopup(); };
    panel.append(search, list); search.addEventListener('input', render); render(); search.focus();
  });
  const renderComments = () => {
    comments.hidden = !commentsOpen; commentButton.setAttribute('aria-expanded', String(commentsOpen)); if (!commentsOpen) return;
    const draft = comments.querySelector('textarea')?.value ?? ''; comments.replaceChildren();
    const heading = document.createElement('strong'); heading.textContent = '代码块评论'; comments.append(heading);
    const entries: Comment[] = node.attrs.comments; if (!entries.length) { const empty = document.createElement('p'); empty.textContent = '暂无评论'; comments.append(empty); }
    entries.forEach(entry => { const row = document.createElement('div'); row.className = 'code-comment'; const text = document.createElement('p'); text.textContent = entry.text; const date = document.createElement('small'); date.textContent = new Date(entry.createdAt).toLocaleString('zh-CN'); const remove = button('删除评论', 'delete', () => { patch({ comments: entries.filter(item => item.id !== entry.id) }); }); row.append(text, date, remove); comments.append(row); });
    const form = document.createElement('form'); const input = document.createElement('textarea'); input.setAttribute('aria-label', '代码评论内容'); input.placeholder = '写下评论…'; input.value = draft;
    const submit = document.createElement('button'); submit.type = 'submit'; submit.textContent = '添加评论'; submit.disabled = !draft.trim(); input.addEventListener('input', () => { submit.disabled = !input.value.trim(); }); form.append(input, submit); form.addEventListener('submit', event => { event.preventDefault(); if (!input.value.trim()) return; const text = input.value.trim(); input.value = ''; patch({ comments: [...node.attrs.comments, { id: crypto.randomUUID(), text, createdAt: Date.now() }] }); comments.querySelector('textarea')?.focus(); }); comments.append(form);
  };
  const measureLines = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(() => {
    if (!dom.isConnected) return; gutter.replaceChildren();
    // DOM ranges use the real highlighted code layout, so wrapped logical lines keep aligned numbers.
    const walker = document.createTreeWalker(contentDOM, NodeFilter.SHOW_TEXT); const texts: Text[] = []; let text = walker.nextNode(); while (text) { texts.push(text as Text); text = walker.nextNode(); }
    const bounds = contentDOM.getBoundingClientRect(); const full = node.textContent; const lineHeight = parseFloat(getComputedStyle(contentDOM).lineHeight); let offset = 0; let line = 1;
    full.split('\n').forEach(value => { const number = document.createElement('span'); number.textContent = `${line++}\n`; let remaining = offset; let top = (line - 2) * lineHeight;
      for (const text of texts) { if (remaining <= text.length) { const range = document.createRange(); range.setStart(text, remaining); range.setEnd(text, Math.min(text.length, remaining + 1)); const rect = range.getClientRects()[0]; if (rect?.height) top = rect.top - bounds.top; break; } remaining -= text.length; }
      number.style.top = `${top}px`; gutter.append(number); offset += value.length + 1;
    });
  }); };
  const update = (next: PMNode) => { if (next.type !== node.type) return false; node = next; localId = next.attrs.id || localId; dom.id = `code-${localId}`; dom.dataset.codeId = localId; dom.classList.toggle('code-wrap', node.attrs.wrap); dom.classList.toggle('code-no-numbers', !node.attrs.lineNumbers); dom.classList.toggle('code-collapsed', node.attrs.collapsed); pre.dataset.params = node.attrs.params;
    const lang = aliases[node.attrs.params] || node.attrs.params; language.textContent = `${codeLanguages.find(([id]) => id === lang)?.[1] || node.attrs.params || '纯文本'} ⌄`; const label = node.attrs.collapsed ? '展开代码' : '折叠代码'; collapse.title = label; collapse.setAttribute('aria-label', label); collapse.setAttribute('aria-pressed', String(node.attrs.collapsed)); commentButton.setAttribute('aria-label', `代码块评论${node.attrs.comments.length ? `（${node.attrs.comments.length}）` : ''}`); renderComments(); measureLines(); return true; };
  const outside = (event: MouseEvent) => { if (popup && !popup.contains(event.target as Node) && !toolbar.contains(event.target as Node)) closePopup(); };
  document.addEventListener('mousedown', outside); window.addEventListener('resize', positionPopup); window.addEventListener('scroll', positionPopup, true);
  const observer = new ResizeObserver(measureLines); observer.observe(contentDOM); update(node);
  return { dom, contentDOM, update, stopEvent: event => !pre.contains(event.target as Node), ignoreMutation: mutation => mutation.type !== 'selection' && !contentDOM.contains(mutation.target), destroy() { closePopup(); clearTimeout(statusTimer); cancelAnimationFrame(frame); observer.disconnect(); document.removeEventListener('mousedown', outside); window.removeEventListener('resize', positionPopup); window.removeEventListener('scroll', positionPopup, true); } };
}
