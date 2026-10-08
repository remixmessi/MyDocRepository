import MiniSearch from 'minisearch';
export interface SearchDocument { id: string; title: string; text: string }
export function tokenize(text: string): string[] {
  const words = [...new Intl.Segmenter('zh', { granularity: 'word' }).segment(text.toLowerCase())].filter(part => part.isWordLike).map(part => part.segment);
  const chinese = text.match(/\p{Script=Han}+/gu) ?? [];
  return [...new Set([...words, ...chinese.flatMap(part => { const chars = [...part]; return [...chars, ...chars.slice(0, -1).map((char, index) => char + chars[index + 1])]; })])];
}
export class DocumentSearch {
  private contents = new Map<string, string>();
  text(id: string) { return this.contents.get(id) ?? ''; }
  private index = new MiniSearch<SearchDocument>({ fields: ['title', 'text'], storeFields: ['title', 'text'], tokenize, searchOptions: { prefix: true, combineWith: 'AND', boost: { title: 4 } } });
  update(doc: SearchDocument) { this.contents.set(doc.id, doc.text); if (this.index.has(doc.id)) this.index.replace(doc); else this.index.add(doc); }
  remove(id: string) { this.contents.delete(id); if (this.index.has(id)) this.index.discard(id); }
  search(query: string, limit = 40) { return this.index.search(query.trim()).slice(0, limit).map(hit => ({ id: hit.id as string, title: hit.title as string, text: hit.text as string })); }
}
export function snippet(text: string, query: string) { const lower = text.toLowerCase(); const words = tokenize(query).sort((a, b) => b.length - a.length); const match = words.map(word => lower.indexOf(word)).find(index => index >= 0) ?? 0; const start = Math.max(0, match - 32); return `${start ? '…' : ''}${text.slice(start, start + 130).replace(/\n/g, ' ')}${text.length > start + 130 ? '…' : ''}`; }
