import type { Node } from 'prosemirror-model';
export interface ReferenceTarget { id: string; title: string }
export interface OutlineEntry { position: number; level: number; text: string }
export function referenceHref(workspaceId: string, docId: string) { return `#${new URLSearchParams({ workspace: workspaceId, doc: docId })}`; }
export function parseReferenceHref(href: string) {
  if (!href.startsWith('#')) return null;
  const params = new URLSearchParams(href.slice(1)); const docId = params.get('doc');
  return docId && !params.has('block') ? { docId, workspaceId: params.get('workspace') || 'local' } : null;
}
export function documentText(node: Node, resolve?: (id: string, workspaceId: string) => string | undefined) {
  return node.textBetween(0, node.content.size, '\n', leaf => ['math_inline','math_block','mermaid'].includes(leaf.type.name)?leaf.attrs.source:['doc_reference', 'doc_embed','surface_ref'].includes(leaf.type.name) ? resolve?.(leaf.attrs.docId, leaf.attrs.workspaceId) || leaf.attrs.label : leaf.type.name === 'web_card' ? `${leaf.attrs.title} ${leaf.attrs.description} ${leaf.attrs.url}` : leaf.type.name === 'attachment' ? leaf.attrs.name : leaf.type.name === 'image' ? leaf.attrs.alt || '' : '');
}
export function documentOutline(doc: Node, resolve?: (id: string, workspaceId: string) => string | undefined): OutlineEntry[] {
  const entries: OutlineEntry[] = [];
  doc.descendants((node, position) => { if (node.type.name === 'heading') entries.push({ position, level: node.attrs.level, text: documentText(node, resolve) || '无标题' }); }); return entries;
}
export interface ReferenceOccurrence { docId: string; workspaceId: string; label: string; position: number; preview: string }
export function documentReferences(doc: Node, resolve?: (id: string, workspaceId: string) => string | undefined): ReferenceOccurrence[] {
  const references: ReferenceOccurrence[] = [];
  doc.descendants((node, position, parent) => { if (['doc_reference', 'doc_embed','surface_ref'].includes(node.type.name)) references.push({ docId: node.attrs.docId, workspaceId: node.attrs.workspaceId, label: node.attrs.label, position, preview: parent ? documentText(parent, resolve).slice(0, 180) : node.attrs.label }); }); return references;
}
