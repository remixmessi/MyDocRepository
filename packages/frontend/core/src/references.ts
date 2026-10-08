import { documentReferences, type ReferenceOccurrence } from '@mydoc/editor-blocks';
export class ReferenceIndex {
  private sources = new Map<string, ReferenceOccurrence[]>();
  update(id: string, doc: Parameters<typeof documentReferences>[0], resolve?: Parameters<typeof documentReferences>[1], extra:ReferenceOccurrence[]=[]) { this.sources.set(id, [...documentReferences(doc, resolve),...extra]); }
  remove(id: string) { this.sources.delete(id); }
  outgoing(id: string) { return this.sources.get(id) ?? []; }
  backlinks(targetId: string, workspaceId: string) {
    return [...this.sources].flatMap(([sourceId, references]) => {
      if (sourceId === targetId) return [];
      const matches = references.filter(reference => reference.docId === targetId && reference.workspaceId === workspaceId);
      return matches.length ? [{ sourceId, count: matches.length, preview: matches[0].preview }] : [];
    });
  }
}
