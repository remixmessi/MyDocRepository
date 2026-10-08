import {typedReference} from './document-mention';
import { Plugin, Selection, TextSelection, type Command, type EditorState } from 'prosemirror-state';
import { setBlockType, wrapIn } from 'prosemirror-commands';
import { wrapInList } from 'prosemirror-schema-list';
import type { EditorView } from 'prosemirror-view';
import { webCardPositionAllowed } from './web-links';
import { documentSchema as schema } from './document-format';

type Item = { label: string; icon: string; search: string; command: Command; action?: () => void };
const expressionItems:Item[]=[
  {label:'标注框',icon:'!',search:'callout info tip warning biaozhu',command:(state,dispatch)=>webCardPositionAllowed(state.selection)&&wrapIn(schema.nodes.callout,{kind:'note'})(state,dispatch)},
  ...(['math_inline','math_block','mermaid'] as const).map((type):Item=>({label:type==='math_inline'?'行内公式':type==='math_block'?'数学公式':'Mermaid 图表',icon:type==='mermaid'?'◇':'∑',search:type==='mermaid'?'mermaid diagram chart tubiao':'math latex formula gongshi',command:(state,dispatch)=>{if(type!=='math_inline'&&!webCardPositionAllowed(state.selection))return false;const source=type==='mermaid'?'flowchart LR\n  A[开始] --> B[完成]':'x^2 + y^2 = z^2';const tr=state.tr.replaceSelectionWith(schema.nodes[type].create({source}));if(type!=='math_inline'){const pos=tr.selection.to;tr.insert(pos,schema.nodes.paragraph.create());tr.setSelection(TextSelection.near(tr.doc.resolve(pos+1)));}dispatch?.(tr.scrollIntoView());return true;}})),
];
const items: Item[] = [
  { label: '文本', icon: 'T', search: 'text paragraph wenben', command: (state, dispatch) => state.selection.$from.parent.type === schema.nodes.paragraph || setBlockType(schema.nodes.paragraph)(state, dispatch) },
  ...Array.from({ length: 6 }, (_, i): Item => ({ label: `标题 ${i + 1}`, icon: `H${i + 1}`, search: `heading h${i + 1} title biaoti`, command: (state, dispatch) => (state.selection.$from.parent.type === schema.nodes.heading && state.selection.$from.parent.attrs.level === i + 1) || setBlockType(schema.nodes.heading, { level: i + 1 })(state, dispatch) })),
  { label: '无序列表', icon: '•', search: 'bullet list wuxuliebiao', command: wrapInList(schema.nodes.bullet_list) },
  { label: '有序列表', icon: '1.', search: 'ordered numbered list youxuliebiao', command: wrapInList(schema.nodes.ordered_list) },
  { label: '待办清单', icon: '☑', search: 'task todo checkbox daiban', command: (state, dispatch) => wrapInList(schema.nodes.bullet_list)(state, tr => {
    const { $from } = tr.selection; for (let depth = $from.depth; depth > 0; depth--) if ($from.node(depth).type === schema.nodes.list_item) { tr.setNodeMarkup($from.before(depth), undefined, { checked: false }); break; }
    dispatch?.(tr);
  }) },
  { label: '引用', icon: '❝', search: 'quote blockquote yinyong', command: wrapIn(schema.nodes.blockquote) },
  { label: '代码块', icon: '</>', search: 'code fence daima', command: setBlockType(schema.nodes.code_block) },
  { label: '分隔线', icon: '—', search: 'divider horizontal rule fengexian', command: (state, dispatch) => {
    const tr = state.tr.replaceSelectionWith(schema.nodes.horizontal_rule.create());
    const pos = tr.selection.to; tr.insert(pos, schema.nodes.paragraph.create()); tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 1))); dispatch?.(tr); return true;
  } },
  { label: '表格', icon: '▦', search: 'table biaoge', command: (state, dispatch) => {
    const rows = Array.from({ length: 3 }, (_, row) => schema.nodes.table_row.create(null, Array.from({ length: 3 }, () => (row === 0 ? schema.nodes.table_header : schema.nodes.table_cell).createAndFill()!)));
    const tr = state.tr.replaceSelectionWith(schema.nodes.table.create(null, rows));
    // The selection after insertion is at the end; place it in the first cell.
    let firstCell: number | undefined; tr.doc.descendants((node, pos) => { if (firstCell === undefined && node.type === schema.nodes.table_header && pos >= state.selection.$from.before()) firstCell = pos + 2; });
    if (firstCell !== undefined) tr.setSelection(TextSelection.create(tr.doc, firstCell)); dispatch?.(tr); return true;
  } },
];
function match(view: EditorView) {
  const { $from, empty } = view.state.selection;
  if (!empty || !$from.parent.isTextblock || $from.parent.type.spec.code || view.composing) return null;
  const text = $from.parent.textBetween(0, $from.parentOffset, '\n', '\ufffc');
  if (typedReference(view.state,view.composing)) return null;
  const found = /(?:^|\s)\/([^/\s]*)$/.exec(text);
  return found ? { from: $from.pos - found[1].length - 1, to: $from.pos, query: found[1] } : null;
}
let slashSerial = 0;
export function documentSlashMenu(onFile?: (kind: 'image' | 'attachment') => void, onWeb?: (kind: 'bookmark' | 'embed') => void, onEmbedDoc?: () => void,onSurface?:()=>void,onFootnote?:()=>void) {
  const menuId = `document-slash-menu-${++slashSerial}`;
  const menuItems: Item[] = [...items,...expressionItems,...(onFootnote?[{label:'脚注',icon:'¹',search:'footnote note jiaozhu',command:()=>true,action:onFootnote}]:[]),...(onSurface?[{label:'引用画布',icon:'▧',search:'surface-ref surface canvas huabu',command:(state:EditorState)=>webCardPositionAllowed(state.selection),action:onSurface}]:[]), ...(onEmbedDoc ? [{ label: '嵌入文档', icon: '▤', search: 'embed-doc embed doc qianruwendang', command: (state: EditorState) => webCardPositionAllowed(state.selection), action: onEmbedDoc }] : []), ...(onWeb ? [{ label: '网页书签', icon: '↗', search: 'bookmark link shuqian', command: (state: EditorState) => webCardPositionAllowed(state.selection), action: () => onWeb('bookmark') }, { label: '嵌入网页', icon: '▣', search: 'embed iframe qianru', command: (state: EditorState) => webCardPositionAllowed(state.selection), action: () => onWeb('embed') }] : []), ...(onFile ? [{ label: '上传图片', icon: '▧', search: 'image photo upload tupian', command: () => true, action: () => onFile('image') }, { label: '添加附件', icon: '↧', search: 'attachment file fujian', command: () => true, action: () => onFile('attachment') }] : [])];
  let active: ReturnType<typeof match> = null; let filtered: Item[] = []; let selected = 0; let dismissed = ''; let popup: HTMLDivElement; let editor: EditorView;
  const signature = () => active ? `${active.from}:${active.to}:${active.query}` : '';
  const choose = (item: Item) => {
    if (!active) return;
    const tr = editor.state.tr.delete(active.from, active.to); const temporary = editor.state.apply(tr);
    const applied = item.command(temporary, result => { result.steps.forEach(step => tr.step(step)); tr.setSelection(Selection.fromJSON(tr.doc, result.selection.toJSON())); });
    if (!applied) return;
    dismissed = ''; active = null; popup.hidden = true; editor.dispatch(tr.scrollIntoView()); editor.focus(); item.action?.();
  };
  const position = () => {
    if (!active || popup.hidden) return;
    const coords = editor.coordsAtPos(active.to); const width = Math.min(280, window.innerWidth - 24);
    popup.style.width = `${width}px`; popup.style.left = `${Math.max(12, Math.min(coords.left, window.innerWidth - width - 12))}px`;
    const height = Math.min(popup.scrollHeight, 340, window.innerHeight - 24);
    popup.style.maxHeight = `${height}px`; popup.style.top = `${coords.bottom + height + 8 < window.innerHeight ? coords.bottom + 8 : Math.max(12, coords.top - height - 8)}px`;
  };
  const render = () => {
    popup.replaceChildren(); popup.hidden = !active || signature() === dismissed || !editor.hasFocus(); if (popup.hidden && editor.dom.getAttribute('aria-controls') && editor.dom.getAttribute('aria-controls') !== popup.id) return;
    editor.dom.setAttribute('aria-expanded', String(!popup.hidden));
    if (popup.hidden) { editor.dom.removeAttribute('aria-activedescendant'); return; }
    editor.dom.setAttribute('aria-controls', popup.id); editor.dom.setAttribute('aria-activedescendant', `${menuId}-option-${selected}`);
    const title = document.createElement('div'); title.className = 'slash-menu-label'; title.textContent = '插入内容'; popup.append(title);
    filtered.forEach((item, index) => {
      const button = document.createElement('button'); button.type = 'button'; button.role = 'option'; button.setAttribute('aria-selected', String(index === selected)); button.id = `${menuId}-option-${index}`;
      const icon = document.createElement('span'); icon.className = 'slash-menu-icon';icon.setAttribute('aria-hidden','true'); icon.textContent = item.icon;
      button.append(icon, document.createTextNode(item.label)); button.addEventListener('mousedown', event => { event.preventDefault(); choose(item); }); popup.append(button);
    });
    if (!filtered.length) { const empty = document.createElement('div'); empty.className = 'slash-menu-label'; empty.textContent = '没有匹配的格式'; popup.append(empty); }
    position(); popup.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
  };
  const close = () => { dismissed = signature(); popup.hidden = true; editor.dom.setAttribute('aria-expanded', 'false'); editor.dom.removeAttribute('aria-activedescendant'); };
  return new Plugin({
    props: {
      handleKeyDown(_view, event) {
        if (!active || popup.hidden || event.isComposing) return false;
        if (event.key === 'Escape') { close(); return true; }
        if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { if (filtered.length) selected = (selected + (event.key === 'ArrowDown' ? 1 : -1) + filtered.length) % filtered.length; render(); return true; }
        if (event.key === 'Enter' && filtered[selected]) { choose(filtered[selected]); return true; }
        return false;
      },
      handleDOMEvents: { blur: () => { popup.hidden = true; return false; }, focus: () => { queueMicrotask(render); return false; } },
    },
    view(view) {
      editor = view; popup = document.createElement('div'); popup.className = 'slash-menu'; popup.role = 'listbox'; popup.id = menuId; popup.setAttribute('aria-label', '插入内容'); popup.hidden = true; document.body.append(popup);
      const outside = (event: MouseEvent) => { if (!popup.contains(event.target as Node) && !editor.dom.contains(event.target as Node)) close(); };
      document.addEventListener('mousedown', outside); window.addEventListener('resize', position); window.addEventListener('scroll', position, true);
      return { update(view) { editor = view; const previous = signature(); active = match(view); if (!active) dismissed = ''; if (signature() !== previous) selected = 0;
        filtered = active ? menuItems.filter(item => `${item.label} ${item.search}`.toLowerCase().includes(active!.query.toLowerCase()) && item.command(view.state.apply(view.state.tr.delete(active!.from, active!.to)))) : []; selected = Math.min(selected, Math.max(0, filtered.length - 1)); render(); },
        destroy() { popup.remove(); document.removeEventListener('mousedown', outside); window.removeEventListener('resize', position); window.removeEventListener('scroll', position, true); } };
    },
  });
}
