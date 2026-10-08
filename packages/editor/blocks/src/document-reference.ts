import {typedReference} from './document-mention';
import { Plugin } from 'prosemirror-state';
import type { EditorView, NodeView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import type { BlockStore } from '@mydoc/editor-framework';
import { documentSchema } from './document-format';
import { referenceHref, type ReferenceTarget } from './document-links';
export interface ReferenceContext { workspaceId: string; documentId: string; documents: ReferenceTarget[]; navigate: (id: string, workspaceId: string, frameId?:string) => void; notify: (message: string) => void }
let menuSerial=0;
export function referenceMenu(context: () => ReferenceContext, store: BlockStore) {
  const menuId=`reference-candidates-${++menuSerial}`;
  let editor: EditorView; let popup: HTMLDivElement; let manual = false; let range: { from: number; to: number; query: string } | null = null;
  let selected = 0; let dismissed = ''; let matches: ReferenceTarget[] = [];
  const signature = () => range ? `${range.from}:${range.to}:${range.query}` : '';
  const typed = (view:EditorView)=>typedReference(view.state,view.composing);
  const close = (focus = false) => { dismissed = signature(); manual = false; range = null; popup.hidden = true; if (editor.dom.getAttribute('aria-controls') === menuId) { editor.dom.removeAttribute('aria-activedescendant'); editor.dom.setAttribute('aria-expanded', 'false'); } if (focus) editor.focus(); };
  const position = () => {
    if (!range || popup.hidden) return; const coords = editor.coordsAtPos(Math.min(range.to, editor.state.doc.content.size)); const width = Math.min(300, innerWidth - 24); popup.style.width = `${width}px`; popup.style.left = `${Math.max(12, Math.min(coords.left, innerWidth - width - 12))}px`;
    const height = Math.min(popup.scrollHeight, 340, innerHeight - 24); popup.style.maxHeight = `${height}px`; popup.style.top = `${coords.bottom + height + 8 < innerHeight ? coords.bottom + 8 : Math.max(12, coords.top - height - 8)}px`;
  };
  const select = (target: ReferenceTarget) => {
    if (!range || !context().documents.some(doc => doc.id === target.id)) return;
    const reference = documentSchema.nodes.doc_reference.create({ docId: target.id, workspaceId: context().workspaceId, label: target.title.trim() || '无标题文档' });
    const { from, to } = range; close(); store.history.stopCapturing(); editor.dispatch(editor.state.tr.replaceWith(from, to, reference).scrollIntoView()); store.history.stopCapturing(); editor.focus();
  };
  const keyboard = (event: KeyboardEvent) => {
    if (!range || popup.hidden || event.isComposing) return false;
    if (event.key === 'Escape') { close(true); return true; }
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { if (matches.length) selected = (selected + (event.key === 'ArrowDown' ? 1 : -1) + matches.length) % matches.length; renderOptions(); return true; }
    if (event.key === 'Enter' && matches[selected]) { select(matches[selected]); return true; } return false;
  };
  const renderOptions = () => {
    const list = popup.querySelector<HTMLDivElement>('.reference-options')!; list.replaceChildren();
    matches = context().documents.filter(doc => doc.id !== context().documentId && doc.title.toLowerCase().includes(range?.query.toLowerCase().trim() || '')).slice(0, 50); selected = Math.min(selected, Math.max(0, matches.length - 1));
    for (const [index, target] of matches.entries()) { const item = document.createElement('button'); item.type = 'button'; item.role = 'option'; item.id = `${menuId}-${index}`; item.setAttribute('aria-selected', String(selected === index)); item.textContent = target.title.trim() || '无标题文档'; item.addEventListener('mousedown', event => { event.preventDefault(); select(target); }); list.append(item); }
    if (!matches.length) { const empty = document.createElement('p'); empty.textContent = '没有匹配文档，请先在当前工作区新建文档。'; list.append(empty); }
    const input = manual ? popup.querySelector('input') : editor.dom; if (input) { input.setAttribute('aria-controls', menuId); input.setAttribute('aria-expanded', 'true'); if (matches[selected]) input.setAttribute('aria-activedescendant', `${menuId}-${selected}`); else input.removeAttribute('aria-activedescendant'); }
    list.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' }); position();
  };
  const render = () => {
    popup.replaceChildren(); if (!range || signature() === dismissed || (!manual && !editor.hasFocus())) { popup.hidden = true; if(editor.dom.getAttribute('aria-controls')===menuId){editor.dom.setAttribute('aria-expanded','false');editor.dom.removeAttribute('aria-activedescendant');} return; } popup.hidden = false;
    if (manual) { const input = document.createElement('input'); input.type = 'search'; input.placeholder = '搜索要引用的文档'; input.setAttribute('aria-label', '搜索引用文档'); input.value = range.query; input.addEventListener('input', () => { if (range) range.query = input.value; selected = 0; renderOptions(); }); input.addEventListener('keydown', event => { if (keyboard(event)) event.preventDefault(); }); popup.append(input); }
    else { const heading = document.createElement('div'); heading.className = 'slash-menu-label'; heading.textContent = '引用文档'; popup.append(heading); }
    const list = document.createElement('div'); list.className = 'reference-options'; list.id = menuId; list.role = 'listbox'; list.setAttribute('aria-label', '候选文档'); popup.append(list); renderOptions(); if (manual) popup.querySelector('input')?.focus();
  };
  const refreshNodes = () => { if (!editor) return; for (const link of editor.dom.querySelectorAll<HTMLAnchorElement>('a[data-doc-reference]')) {
    const id = link.dataset.docReference!; const workspace = link.dataset.workspace!; const target = workspace === context().workspaceId ? context().documents.find(doc => doc.id === id) : undefined; const missing = workspace === context().workspaceId && !target;
    link.textContent = `↗ ${target?.title.trim() || link.dataset.referenceLabel || '无标题文档'}${missing ? '（已删除）' : ''}`; link.classList.toggle('missing-reference', missing); link.setAttribute('aria-disabled', String(missing)); link.title = missing ? '引用的文档已删除' : '打开引用文档';
  } };
  const plugin = new Plugin({
    props: { handleKeyDown: (_view, event) => keyboard(event), handleDOMEvents: { blur: () => { if (!manual) close(); return false; }, compositionstart:()=>{popup.hidden=true;return false;},compositionend:()=>{queueMicrotask(()=>{range=typed(editor);render();});return false;},focus:()=>{queueMicrotask(()=>{if(!manual){range=typed(editor);render();}});return false;} } },
    view(view) {
      editor = view; popup = document.createElement('div'); popup.className = 'reference-popup'; popup.role = 'dialog'; popup.setAttribute('aria-label', '引用文档'); popup.hidden = true; document.body.append(popup);
      const outside = (event: MouseEvent) => { if (!popup.contains(event.target as Node) && (!editor.dom.contains(event.target as Node) || manual)) close(); };
      document.addEventListener('mousedown', outside); window.addEventListener('resize', position); window.addEventListener('scroll', position, true);
      queueMicrotask(refreshNodes);
      return { update(view) { editor = view; refreshNodes(); if (manual) { position(); return; } const previous = signature(); range = typed(view); if (!range) dismissed = ''; if (signature() !== previous) selected = 0; render(); }, destroy() { popup.remove(); document.removeEventListener('mousedown', outside); window.removeEventListener('resize', position); window.removeEventListener('scroll', position, true); } };
    },
  });
  const nodeView = (node: PMNode): NodeView => { const dom = document.createElement('a'); dom.className = 'doc-reference'; dom.contentEditable = 'false';
    const update = (next: PMNode) => { if (next.type !== node.type) return false; node = next; dom.href = referenceHref(next.attrs.workspaceId, next.attrs.docId); dom.dataset.docReference = next.attrs.docId; dom.dataset.workspace = next.attrs.workspaceId; dom.dataset.referenceLabel = next.attrs.label; dom.textContent = `↗ ${next.attrs.label}`; return true; };
    dom.addEventListener('click', event => { if (event.metaKey || event.ctrlKey || event.shiftKey) return; event.preventDefault(); if (dom.classList.contains('missing-reference')) context().notify('引用的文档已删除，正文内容已保留。'); else context().navigate(node.attrs.docId, node.attrs.workspaceId); }); update(node);
    return { dom, update, stopEvent: event => event.type === 'click', ignoreMutation: () => true };
  };
  return { plugin, nodeView, refresh() { refreshNodes(); if (range && !popup.hidden) renderOptions(); }, open() {
    if (editor.state.selection.$from.parent.type.spec.code || editor.state.selection.$from.marks().some(mark=>mark.type.spec.code) || !editor.state.selection.$from.parent.isTextblock) { context().notify('请在正文文本中选择引用位置。'); return; }
    dismissed = ''; manual = true; range = { from: editor.state.selection.from, to: editor.state.selection.to, query: '' }; selected = 0; render();
  } };
}
