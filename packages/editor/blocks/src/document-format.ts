import {footnoteNodes,footnoteMarkdown,footnoteSerializers} from './footnote-format';
import{expressionNodes,expressionMarkdown,expressionSerializers}from'./expression-format';
import {validTemplate,type BoardTemplate} from './whiteboard/templates';
import {parseSurfaceRef,type SurfaceRefAttrs} from './surface-ref-format';
import {validSurfaceElement,type SurfaceElement} from './whiteboard/model';
import MarkdownIt from 'markdown-it';
import { Fragment, Schema, type Node } from 'prosemirror-model';
import { schema as basic, defaultMarkdownParser, defaultMarkdownSerializer, MarkdownParser, MarkdownSerializer } from 'prosemirror-markdown';
import { tableNodes } from 'prosemirror-tables';
import { prosemirrorToYXmlFragment, yXmlFragmentToProsemirrorJSON } from 'y-prosemirror';
import type { BlockStore } from '@mydoc/editor-framework';
import { parseDocEmbed, type DocEmbedAttrs } from './embed-format';
import { parseWebCardMarkdown, webCardMarkdown } from './web-links';
import { documentText, parseReferenceHref, referenceHref } from './document-links';

const listItem = basic.spec.nodes.get('list_item')!;
const nodes = basic.spec.nodes.append(footnoteNodes).append(expressionNodes).addToEnd('surface_ref',{group:'block',atom:true,attrs:{docId:{},workspaceId:{default:'local'},frameId:{},region:{default:null},label:{default:'画布引用'},collapsed:{default:false}},parseDOM:[{tag:'div[data-surface-ref]',getAttrs:element=>parseSurfaceRef((element as HTMLElement).dataset.surfaceRef||'')||false}],toDOM:node=>['div',{'data-surface-ref':JSON.stringify(node.attrs)},['a',{href:referenceHref(node.attrs.workspaceId,node.attrs.docId)+'&frame='+encodeURIComponent(node.attrs.frameId)},node.attrs.label]]}).addToEnd('doc_embed', { group: 'block', atom: true, attrs: { docId: {}, workspaceId: { default: 'local' }, label: { default: '无标题文档' }, collapsed: { default: false } }, parseDOM: [{ tag: 'div[data-doc-embed]', getAttrs: element => parseDocEmbed((element as HTMLElement).dataset.docEmbed || '') || false }], toDOM: node => ['div', { 'data-doc-embed': JSON.stringify(node.attrs) }, ['a', { href: referenceHref(node.attrs.workspaceId, node.attrs.docId) }, node.attrs.label]] }).addToEnd('web_card', { group: 'block', atom: true, attrs: { id: { default: null }, kind: { default: 'bookmark' }, url: { default: '' }, title: { default: '' }, description: { default: '' } }, parseDOM: [{ tag: 'div[data-web-card]', getAttrs: element => parseWebCardMarkdown((element as HTMLElement).dataset.webCard || '') || false }], toDOM: node => ['div', { 'data-web-card': webCardMarkdown(node.attrs as import('./web-links').WebCardAttrs) }, ['a', { href: node.attrs.url }, node.attrs.title || node.attrs.url]] }).addToEnd('attachment', { inline: true, group: 'inline', atom: true, attrs: { assetId: {}, name: { default: '附件' }, size: { default: 0 }, mime: { default: '' } }, parseDOM: [{ tag: 'a[data-attachment]', getAttrs: element => ({ assetId: (element as HTMLElement).dataset.attachment, name: element.textContent || '附件' }) }], toDOM: node => ['a', { href: `asset:${node.attrs.assetId}`, 'data-attachment': node.attrs.assetId }, node.attrs.name] }).addToEnd('doc_reference', { inline: true, group: 'inline', atom: true, attrs: { docId: {}, workspaceId: { default: 'local' }, label: { default: '无标题文档' } },
  parseDOM: [{ tag: 'a[data-doc-reference]', getAttrs: element => ({ docId: (element as HTMLElement).dataset.docReference, workspaceId: (element as HTMLElement).dataset.workspace || 'local', label: element.textContent || '无标题文档' }) }],
  toDOM: node => ['a', { href: referenceHref(node.attrs.workspaceId, node.attrs.docId), 'data-doc-reference': node.attrs.docId, 'data-workspace': node.attrs.workspaceId }, node.attrs.label],
}).update('code_block', { ...basic.spec.nodes.get('code_block')!, attrs: { params: { default: '' }, id: { default: null }, wrap: { default: false }, lineNumbers: { default: true }, collapsed: { default: false }, comments: { default: [] } } }).update('heading', { ...basic.spec.nodes.get('heading')!, content: 'inline*' }).update('list_item', {
  ...listItem, content: 'paragraph block*', attrs: { checked: { default: null } },
  parseDOM: [{ tag: 'li', getAttrs: element => ({ checked: (element as HTMLElement).hasAttribute('data-checked') ? (element as HTMLElement).getAttribute('data-checked') === 'true' : null }) }],
  toDOM: node => ['li', node.attrs.checked === null ? {} : { 'data-checked': String(node.attrs.checked) }, 0],
}).append(tableNodes({ tableGroup: 'block', cellContent: 'block+', cellAttributes: { align: { default: null, getFromDOM: element => element.style.textAlign || null, setDOMAttr: (value, attrs) => { if (['left', 'center', 'right'].includes(value as string)) attrs.style = `text-align: ${value}`; } } } }));
export const documentSchema = new Schema({ nodes, marks: basic.spec.marks.addToEnd('strike', { parseDOM: [{ tag: 's' }, { tag: 'del' }], toDOM: () => ['s', 0] }) });
const tokenizer = MarkdownIt('commonmark', { html: false, linkify: false }).enable(['table', 'strikethrough']);
footnoteMarkdown(tokenizer);
expressionMarkdown(tokenizer);
tokenizer.core.ruler.after('block','surface-refs',state=>{for(const token of state.tokens){if(token.type!=='fence'||token.info.trim()!=='mydoc-surface-ref')continue;const attrs=parseSurfaceRef(token.content);if(!attrs)continue;token.type='surface_ref';token.meta={...attrs};token.content='';}});
tokenizer.core.ruler.after('block', 'doc-embeds', state => { for (const token of state.tokens) { if (token.type !== 'fence' || token.info.trim() !== 'mydoc-embed-doc') continue; const attrs = parseDocEmbed(token.content); if (!attrs) continue; token.type = 'doc_embed'; token.meta = { ...attrs }; token.content = ''; } });
tokenizer.core.ruler.after('block', 'web-cards', state => { for (const token of state.tokens) { if (token.type !== 'fence' || token.info.trim() !== 'mydoc-web-card') continue; const attrs = parseWebCardMarkdown(token.content); if (!attrs) continue; token.type = 'web_card'; token.meta = { ...attrs }; token.content = ''; } });
tokenizer.core.ruler.after('inline', 'document-references', state => {
  for (const token of state.tokens) {
    if (token.type !== 'inline' || !token.children) continue;
    for (let index = 0; index < token.children.length; index++) {
      const open = token.children[index]; if (open.type !== 'link_open') continue;
      const href = String(open.attrGet('href') || ''); const assetId = href.startsWith('asset:') ? href.slice(6) : null; const target = parseReferenceHref(href); if (!target && !assetId) continue;
      const end = token.children.findIndex((child, offset) => offset > index && child.type === 'link_close'); if (end < 0) continue;
      const reference = new state.Token(assetId ? 'attachment' : 'doc_reference', '', 0); if (assetId) reference.attrSet('assetId', assetId); else { reference.attrSet('docId', target!.docId); reference.attrSet('workspaceId', target!.workspaceId); }
      reference.attrSet('label', token.children.slice(index + 1, end).map(child => child.content).join('') || '无标题文档'); token.children.splice(index, end - index + 1, reference);
    }
  }
});
tokenizer.core.ruler.after('inline', 'task-items', state => {
  const withCells: typeof state.tokens = [];
  for (const token of state.tokens) {
    if (token.type === 'th_close' || token.type === 'td_close') withCells.push(new state.Token('paragraph_close', 'p', -1));
    withCells.push(token);
    if (token.type === 'th_open' || token.type === 'td_open') withCells.push(new state.Token('paragraph_open', 'p', 1));
  }
  state.tokens = withCells;
  for (let index = 0; index < state.tokens.length; index++) {
    const item = state.tokens[index]; if (item.type !== 'list_item_open') continue;
    const inline = state.tokens[index + 2]; if (inline?.type !== 'inline') continue;
    const match = /^\[([ xX])\] /.exec(inline.content); if (!match) continue;
    item.attrSet('checked', match[1].toLowerCase() === 'x' ? 'true' : 'false');
    inline.content = inline.content.slice(match[0].length);
    if (inline.children?.[0]?.type === 'text') inline.children[0].content = inline.children[0].content.slice(match[0].length);
  }
});
export const markdownParser = new MarkdownParser(documentSchema, tokenizer, {
  ...defaultMarkdownParser.tokens,footnote_ref:{node:'footnote_reference',getAttrs:token=>({label:token.meta.label})},footnote_definition:{block:'footnote_definition',getAttrs:token=>({label:token.attrGet('label')})},callout:{block:'callout',getAttrs:token=>({kind:token.attrGet('kind')})},math_inline:{node:'math_inline',getAttrs:token=>token.meta},math_block:{node:'math_block',getAttrs:token=>token.meta},mermaid:{node:'mermaid',getAttrs:token=>token.meta}, surface_ref:{node:'surface_ref',getAttrs:token=>token.meta}, doc_embed: { node: 'doc_embed', getAttrs: token => token.meta }, web_card: { node: 'web_card', getAttrs: token => token.meta }, attachment: { node: 'attachment', getAttrs: token => ({ assetId: token.attrGet('assetId'), name: token.attrGet('label') }) }, doc_reference: { node: 'doc_reference', getAttrs: token => ({ docId: token.attrGet('docId'), workspaceId: token.attrGet('workspaceId'), label: token.attrGet('label') }) },
  list_item: { block: 'list_item', getAttrs: token => ({ checked: token.attrGet('checked') === null ? null : token.attrGet('checked') === 'true' }) },
  s: { mark: 'strike' }, table: { block: 'table' }, thead: { ignore: true }, tbody: { ignore: true },
  tr: { block: 'table_row' }, th: { block: 'table_header', getAttrs: token => ({ align: token.attrGet('style')?.replace('text-align:', '').trim() ?? null }) }, td: { block: 'table_cell', getAttrs: token => ({ align: token.attrGet('style')?.replace('text-align:', '').trim() ?? null }) },
});
const inlineSerializer = new MarkdownSerializer({ ...defaultMarkdownSerializer.nodes,...footnoteSerializers,...expressionSerializers, surface_ref(state, node) { state.write('```mydoc-surface-ref\n'+JSON.stringify(node.attrs as SurfaceRefAttrs)+'\n```');state.closeBlock(node); }, doc_embed(state, node) { state.write('```mydoc-embed-doc\n' + JSON.stringify(node.attrs as DocEmbedAttrs) + '\n```'); state.closeBlock(node); }, web_card(state, node) { state.write('```mydoc-web-card\n' + webCardMarkdown(node.attrs as import('./web-links').WebCardAttrs) + '\n```'); state.closeBlock(node); }, attachment(state, node) { state.write('['); state.text(node.attrs.name); state.write(`](asset:${node.attrs.assetId})`); }, doc_reference(state, node) { state.write('['); state.text(node.attrs.label); state.write(`](${referenceHref(node.attrs.workspaceId, node.attrs.docId)})`); } }, { ...defaultMarkdownSerializer.marks, strike: { open: '~~', close: '~~', mixable: true, expelEnclosingWhitespace: true } },{escapeExtraCharacters:/\$/g});
export const markdownSerializer = new MarkdownSerializer({
  ...defaultMarkdownSerializer.nodes,...footnoteSerializers,...expressionSerializers,
  surface_ref(state, node) { state.write('```mydoc-surface-ref\n'+JSON.stringify(node.attrs as SurfaceRefAttrs)+'\n```');state.closeBlock(node); }, doc_embed(state, node) { state.write('```mydoc-embed-doc\n' + JSON.stringify(node.attrs as DocEmbedAttrs) + '\n```'); state.closeBlock(node); },
  web_card(state, node) { state.write('```mydoc-web-card\n' + webCardMarkdown(node.attrs as import('./web-links').WebCardAttrs) + '\n```'); state.closeBlock(node); },
  attachment(state, node) { state.write('['); state.text(node.attrs.name); state.write(`](asset:${node.attrs.assetId})`); },
  doc_reference(state, node) { state.write('['); state.text(node.attrs.label); state.write(`](${referenceHref(node.attrs.workspaceId, node.attrs.docId)})`); },
  list_item(state, node) { if (node.attrs.checked !== null) state.write(node.attrs.checked ? '[x] ' : '[ ] '); state.renderContent(node); },
  table(state, node) {
    const rows: string[][] = [];
    node.forEach(row => { const cells: string[] = []; row.forEach(cell => cells.push(inlineSerializer.serialize(cell).replace(/\n+/g, ' ').replace(/\|/g, '\\|'))); rows.push(cells); });
    rows.forEach((cells, index) => { state.write(`| ${cells.join(' | ')} |\n`); if (!index) state.write(`| ${cells.map((_, column) => ({ left: ':---', center: ':---:', right: '---:' }[node.firstChild!.child(column).attrs.align as string] ?? '---')).join(' | ')} |\n`); });
    state.closeBlock(node);
  },
}, { ...defaultMarkdownSerializer.marks, strike: { open: '~~', close: '~~', mixable: true, expelEnclosingWhitespace: true } },{escapeExtraCharacters:/\$/g});

/** 旧块树只作为备份；新正文存入一个 Y.XmlFragment，迁移仅执行一次。 */
export function migrateDocument(store: BlockStore) {
  const fragment = store.doc.getXmlFragment('document');
  store.history.addToScope(fragment);
  if (store.meta.get('format') === 'markdown-document-v1' || fragment.length) return fragment;
  const paragraph = (id: string) => {
    const json = yXmlFragmentToProsemirrorJSON(store.fragment(id)) as { content?: { content?: Record<string, unknown>[] }[] };
    const inline = json.content?.flatMap((node, index) => [...(index ? [{ type: 'hard_break' }] : []), ...(node.content ?? [])]) ?? [];
    return documentSchema.nodeFromJSON({ type: 'paragraph', content: inline });
  };
  const convert = (ids: string[]): Node[] => {
    const result: Node[] = [];
    for (let index = 0; index < ids.length; index++) {
      const id = ids[index];
      const flavour = store.flavour(id); const props = store.props(id);
      if (flavour === 'note') { result.push(...convert(store.children(id))); continue; }
      if (flavour === 'list') {
        const name = props.type === 'ordered' ? 'ordered_list' : 'bullet_list';
        const items: Node[] = [];
        const group = [id];
        while (index + 1 < ids.length && store.flavour(ids[index + 1]) === 'list' && store.props(ids[index + 1]).type === props.type) group.push(ids[++index]);
        for (const item of group) {
          const content = [paragraph(item), ...convert(store.children(item))];
          items.push(documentSchema.node('list_item', { checked: props.type === 'todo' ? store.props(item).checked : null }, content));
        }
        result.push(documentSchema.node(name, { tight: true }, items));
      } else if (flavour === 'divider') result.push(documentSchema.node('horizontal_rule'));
      else if (flavour === 'heading') result.push(documentSchema.node('heading', { level: props.level }, paragraph(id).content));
      else if (flavour === 'quote') result.push(documentSchema.node('blockquote', null, paragraph(id)));
      else if (flavour === 'code') {
        const text = paragraph(id).textBetween(0, paragraph(id).content.size, '\n', '\n');
        result.push(documentSchema.node('code_block', { params: props.language === 'plain' ? '' : props.language }, text ? documentSchema.text(text) : undefined));
      } else result.push(paragraph(id));
    }
    return result;
  };
  // 深拷贝列表后按游标处理，迁移失败不会删除原块树。
  const content = convert(store.children());
  const doc = documentSchema.node('doc', null, content.length ? content : documentSchema.node('paragraph'));
  store.doc.transact(() => { prosemirrorToYXmlFragment(doc, fragment); store.meta.set('format', 'markdown-document-v1'); }, 'markdown-migration-v1');
  return fragment;
}

/** Keep local code settings and comments when editing the Markdown representation. */
export function preserveCodeMetadata(previous: Node, parsed: Node): Node {
  const existing: Node[] = []; previous.descendants(node => { if (node.type.name === 'code_block') existing.push(node); });
  const used = new Set<Node>(); let index = 0;
  const visit = (node: Node): Node => {
    if (node.type.name === 'code_block') {
      const old = existing.find(candidate => !used.has(candidate) && candidate.textContent === node.textContent && candidate.attrs.params === node.attrs.params) ?? (used.has(existing[index]) ? undefined : existing[index]); index++;
      if (!old) return node; used.add(old); return node.type.create({ ...old.attrs, params: node.attrs.params }, node.content, node.marks);
    }
    if (node.isLeaf) return node;
    const children: Node[] = []; node.forEach(child => children.push(visit(child))); return node.copy(Fragment.fromArray(children));
  };
  return visit(parsed);
}

export function readDocument(store: BlockStore) {
  return documentSchema.nodeFromJSON(yXmlFragmentToProsemirrorJSON(migrateDocument(store)));
}
export function writeMarkdown(store: BlockStore, text: string) {
  const parsed = markdownParser.parse(text.replace(/^\uFEFF/, ''));
  writeDocument(store,parsed);
}
export function writeDocument(store:BlockStore,parsed:Node){
  const fragment=migrateDocument(store);store.doc.transact(()=>prosemirrorToYXmlFragment(parsed,fragment),'document-import');store.history.clear();
}

export function remapReferences(store: BlockStore, fromWorkspace: string, toWorkspace: string, ids: Map<string, string>, onlyMapped = false) {
  const doc = readDocument(store);
  const visit = (node: Node): Node => {
    if (['doc_reference', 'doc_embed','surface_ref'].includes(node.type.name) && node.attrs.workspaceId === fromWorkspace && (!onlyMapped || ids.has(node.attrs.docId))) return node.type.create({ ...node.attrs, workspaceId: toWorkspace, docId: ids.get(node.attrs.docId) || node.attrs.docId }, null, node.marks);
    if (node.isLeaf) return node; const children: Node[] = []; node.forEach(child => children.push(visit(child))); return node.copy(Fragment.fromArray(children));
  };
  const remapElement=(s:SurfaceElement):SurfaceElement=>{if(s.kind==='doc-card'&&s.workspaceId===fromWorkspace&&(!onlyMapped||ids.has(s.docId)))return {...s,workspaceId:toWorkspace,docId:ids.get(s.docId)||s.docId};if(s.kind==='note'){const original=markdownParser.parse(s.markdown),next=visit(original);if(!next.eq(original))return {...s,markdown:markdownSerializer.serialize(next),text:documentText(next).slice(0,2000)};}return s;};
  const surface=store.doc.getMap<SurfaceElement>('surface');store.doc.transact(()=>{for(const [id,s] of surface)if(validSurfaceElement(s)){const next=remapElement(s);if(next!==s)surface.set(id,next);}},'restore-reference-map');
  const templates=store.doc.getMap<BoardTemplate>('whiteboardTemplates');store.doc.transact(()=>{for(const [id,template] of templates)if(validTemplate(template))templates.set(id,{...template,elements:template.elements.map(remapElement)});},'restore-reference-map');const next = visit(doc); if (!next.eq(doc)) store.doc.transact(() => prosemirrorToYXmlFragment(next, migrateDocument(store)), 'restore-reference-map');
}

export function documentAssets(doc: Node) { const ids = new Set<string>(); doc.descendants(node => { if (node.type.name === 'attachment') ids.add(node.attrs.assetId); if (node.type.name === 'image' && String(node.attrs.src).startsWith('asset:')) ids.add(node.attrs.src.slice(6)); }); return ids; }


export function remapAssets(store: BlockStore, ids: Map<string, string>) {
  const visit = (node: Node): Node => {
    if (node.type.name === 'attachment' && ids.has(node.attrs.assetId)) return node.type.create({...node.attrs,assetId:ids.get(node.attrs.assetId)},null,node.marks);
    if (node.type.name === 'image' && String(node.attrs.src).startsWith('asset:') && ids.has(node.attrs.src.slice(6))) return node.type.create({...node.attrs,src:'asset:'+ids.get(node.attrs.src.slice(6))},null,node.marks);
    if(node.isLeaf)return node;const children:Node[]=[];node.forEach(child=>children.push(visit(child)));return node.copy(Fragment.fromArray(children));
  };
  const mapElement=(s:SurfaceElement):SurfaceElement=>{if(s.kind==='image'&&ids.has(s.assetId))return {...s,assetId:ids.get(s.assetId)!};if(s.kind==='note'){const original=markdownParser.parse(s.markdown),next=visit(original);if(!next.eq(original))return {...s,markdown:markdownSerializer.serialize(next)};}return s;};
  store.doc.transact(()=>{const surface=store.doc.getMap<SurfaceElement>('surface');for(const [id,s]of surface)surface.set(id,mapElement(s));const templates=store.doc.getMap<BoardTemplate>('whiteboardTemplates');for(const[id,t]of templates)templates.set(id,{...t,elements:t.elements.map(mapElement)});const original=readDocument(store),next=visit(original);if(!next.eq(original))prosemirrorToYXmlFragment(next,migrateDocument(store));},'template-assets');
}
