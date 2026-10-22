import { Y, LOCAL_ORIGIN } from '@mydoc/crdt';
import type { DocumentMeta } from './workspace';
export const TAG_COLORS = [
  { id: 'default', label: '默认' }, { id: 'red', label: '红色' },
  { id: 'orange', label: '橙色' }, { id: 'yellow', label: '黄色' },
  { id: 'green', label: '绿色' }, { id: 'blue', label: '蓝色' },
  { id: 'purple', label: '紫色' }, { id: 'pink', label: '粉色' },
] as const;
export type TagColor = typeof TAG_COLORS[number]['id'];
export const validTagColor = (value: unknown): value is TagColor => TAG_COLORS.some(color => color.id === value);
interface NamedMeta { id: string; name: string; createdAt: number }
export interface TagMeta extends NamedMeta { color?: TagColor }
export interface CollectionRules { tagIds: string[]; tagMatch: 'all' | 'any'; keyword: string; favorite: 'any' | 'yes' | 'no'; paused?: boolean }
export interface CollectionMeta { id: string; name: string; createdAt: number; mode: 'manual' | 'rules'; documentIds: string[]; rules: CollectionRules }
export const emptyRules = (): CollectionRules => ({ tagIds: [], tagMatch: 'all', keyword: '', favorite: 'any' });
const record = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value);
export const uniqueIds = (value: unknown): value is string[] => Array.isArray(value) && value.length <= 10000 && value.every(id => typeof id === 'string' && !!id) && new Set(value).size === value.length;
function validNamedMeta(value: unknown): value is NamedMeta { return record(value) && typeof value.id === 'string' && !!value.id && typeof value.name === 'string' && !!value.name.trim() && value.name.length <= 160 && typeof value.createdAt === 'number' && Number.isFinite(value.createdAt); }
export function validTag(value: unknown): value is TagMeta { return validNamedMeta(value) && record(value) && (value.color === undefined || validTagColor(value.color)); }
export function validCollection(value: unknown): value is CollectionMeta {
  return validNamedMeta(value) && record(value) && ['manual', 'rules'].includes(String(value.mode)) && uniqueIds(value.documentIds) && record(value.rules) && uniqueIds(value.rules.tagIds) && ['all', 'any'].includes(String(value.rules.tagMatch)) && typeof value.rules.keyword === 'string' && value.rules.keyword.length <= 160 && ['any', 'yes', 'no'].includes(String(value.rules.favorite)) && (value.rules.paused === undefined || typeof value.rules.paused === 'boolean');
}
function nameFor(name: string, items: TagMeta[], id?: string) {
  const clean = name.trim(); if (!clean || clean.length > 160) throw new Error('名称需为 1–160 个字符');
  if (items.some(item => item.id !== id && item.name.toLocaleLowerCase() === clean.toLocaleLowerCase())) throw new Error('名称已存在'); return clean;
}
export class DocumentOrganization {
  readonly tags: Y.Map<TagMeta>; readonly collections: Y.Map<CollectionMeta>;
  constructor(private doc: Y.Doc, private documents: {get:(id:string)=>DocumentMeta|undefined;has:(id:string)=>boolean;values:()=>IterableIterator<DocumentMeta>;set:(id:string,meta:DocumentMeta)=>unknown}, private text: (id: string) => string,private assertWritable:()=>void=()=>{}) { this.tags = doc.getMap('tags'); this.collections = doc.getMap('collections'); }
  tagList() { return [...this.tags.values()].sort((a,b) => a.name.localeCompare(b.name, 'zh-CN') || a.id.localeCompare(b.id)); }
  collectionList() { return [...this.collections.values()].sort((a,b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id)); }
  saveTag(name: string, id?: string, color?: TagColor) { this.assertWritable();
    if (id && !this.tags.has(id)) throw new Error('标签已删除'); const clean = nameFor(name, this.tagList(), id); const key = id ?? crypto.randomUUID();
    const selectedColor = color === undefined ? this.tags.get(key)?.color : color;
    if (selectedColor !== undefined && !validTagColor(selectedColor)) throw new Error('标签颜色无效');
    this.doc.transact(() => this.tags.set(key, { id: key, name: clean, createdAt: this.tags.get(key)?.createdAt ?? Date.now(), ...(selectedColor !== undefined ? { color: selectedColor } : {}) }), LOCAL_ORIGIN); return key;
  }
  deleteTag(id: string) { this.assertWritable();
    this.doc.transact(() => { this.tags.delete(id); for (const meta of this.documents.values()) if (meta.tagIds?.includes(id)) this.documents.set(meta.id, { ...meta, tagIds: meta.tagIds.filter(tag => tag !== id) });
      for (const collection of this.collections.values()) if (collection.rules.tagIds.includes(id)) this.collections.set(collection.id, { ...collection, rules: { ...collection.rules, tagIds: collection.rules.tagIds.filter(tag => tag !== id), paused: collection.mode === 'rules' } });
    }, LOCAL_ORIGIN);
  }
  setFavorite(id: string, favorite: boolean) { this.assertWritable(); const meta = this.documents.get(id); if (!meta) throw new Error('文档已删除'); this.doc.transact(() => this.documents.set(id, { ...meta, favorite }), LOCAL_ORIGIN); }
  setDocumentTags(id: string, tagIds: string[]) { this.assertWritable();
    const meta = this.documents.get(id); if (!meta) throw new Error('文档已删除'); if (!uniqueIds(tagIds) || tagIds.some(tag => !this.tags.has(tag))) throw new Error('标签不存在');
    this.doc.transact(() => this.documents.set(id, { ...meta, tagIds: [...tagIds] }), LOCAL_ORIGIN);
  }
  saveCollection(input: Omit<CollectionMeta, 'id' | 'createdAt'>, id?: string) { this.assertWritable();
    if (id && !this.collections.has(id)) throw new Error('集合已删除'); const key = id ?? crypto.randomUUID(); const value: CollectionMeta = { ...input, name: nameFor(input.name, this.collectionList(), id), id: key, createdAt: this.collections.get(key)?.createdAt ?? Date.now(), documentIds: [...input.documentIds], rules: { ...input.rules, tagIds: [...input.rules.tagIds], keyword: input.rules.keyword.trim() } };
    if (!validCollection(value) || value.documentIds.some(doc => !this.documents.has(doc)) || value.rules.tagIds.some(tag => !this.tags.has(tag))) throw new Error('集合包含无效的文档或规则');
    if (value.mode === 'rules' && !value.rules.tagIds.length && !value.rules.keyword && value.rules.favorite === 'any') throw new Error('规则集合至少需要一个筛选条件');
    this.doc.transact(() => this.collections.set(key, value), LOCAL_ORIGIN); return key;
  }
  deleteCollection(id: string) { this.assertWritable(); this.doc.transact(() => this.collections.delete(id), LOCAL_ORIGIN); }
  setMembership(collectionId: string, documentId: string, include: boolean) { this.assertWritable();
    const value = this.collections.get(collectionId); if (!value || !this.documents.has(documentId)) throw new Error('文档或集合已删除'); if (value.mode !== 'manual') throw new Error('规则集合自动归集，请修改规则');
    const documentIds = value.documentIds.filter(id => id !== documentId); if (include) documentIds.push(documentId);
    this.doc.transact(() => this.collections.set(collectionId, { ...value, documentIds }), LOCAL_ORIGIN);
  }
  assignDocument(id: string, tagIds: string[], collectionIds: string[]) { this.assertWritable();
    if (!this.documents.has(id) || !uniqueIds(tagIds) || tagIds.some(tag => !this.tags.has(tag)) || !uniqueIds(collectionIds) || collectionIds.some(key => this.collections.get(key)?.mode !== 'manual')) throw new Error('文档、标签或集合已变化，请重新打开设置');
    this.doc.transact(() => { this.setDocumentTags(id, tagIds); for (const c of [...this.collections.values()]) if (c.mode === 'manual') this.setMembership(c.id, id, collectionIds.includes(c.id)); }, LOCAL_ORIGIN);
  }
  removeDocument(id: string) { this.assertWritable(); for (const value of [...this.collections.values()]) if (value.documentIds.includes(id)) this.collections.set(value.id, { ...value, documentIds: value.documentIds.filter(doc => doc !== id) }); }
  matches(meta: DocumentMeta, value: CollectionMeta) {
    if (value.mode === 'manual') return value.documentIds.includes(meta.id);
    const r = value.rules; if (r.paused || (!r.tagIds.length && !r.keyword && r.favorite === 'any')) return false;
    if (r.tagIds.length && !(r.tagMatch === 'all' ? r.tagIds.every(id => meta.tagIds?.includes(id)) : r.tagIds.some(id => meta.tagIds?.includes(id)))) return false;
    if (r.favorite !== 'any' && !!meta.favorite !== (r.favorite === 'yes')) return false;
    return !r.keyword || `${meta.title}\n${this.text(meta.id)}`.toLocaleLowerCase().includes(r.keyword.toLocaleLowerCase());
  }
  members(id: string) { const value = this.collections.get(id); return value ? [...this.documents.values()].filter(meta => this.matches(meta, value)) : []; }
}
