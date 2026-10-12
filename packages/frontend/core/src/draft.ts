/**
 * 草稿文件解析、恢复预览与恢复。
 *
 * 草稿由状态区的“导出草稿”生成，有两种格式：
 * - `mydoc-workspace-draft`：工作空间草稿，含已打开文档的 Yjs 状态与 Markdown。
 * - `mydoc-document-draft`：单文档草稿，只有一份 Yjs 状态。
 *
 * 草稿不含附件字节，也不含服务器权限信息。恢复一律在个人工作区创建新文档并分配新 ID：
 * 不复用草稿中的文档 ID，因此不会绕过服务端 tombstone，也不会重新获得原文档访问权。
 */
import {Y} from '@mydoc/crdt';
import {BlockStore} from '@mydoc/editor-framework';
import {readDocument,documentAssets,markdownSerializer,validSurface,type SurfaceElement} from '@mydoc/editor-blocks';
import type {LocalWorkspace} from './workspace';

/** 草稿文件大小上限。草稿只含正文与白板 CRDT，不含附件字节。 */
export const MAX_DRAFT_SIZE = 50 * 1024 * 1024;
const MAX_DRAFT_DOCUMENTS = 1000;
const SUMMARY_LENGTH = 200;

export interface DraftDocument {
  /** 草稿中的原文档 ID，仅用于展示与去重，恢复时不复用。 */
  sourceId: string;
  title: string;
  summary: string;
  /** 白板元素数量，0 表示该文档没有白板内容。 */
  surfaceElements: number;
  /** 正文与白板引用但草稿未携带字节的附件 ID。 */
  missingAssets: string[];
  update: Uint8Array;
}
export interface DraftPreview {
  kind: 'workspace' | 'document';
  /** 草稿来源工作空间 ID，单文档草稿没有。 */
  sourceWorkspaceId?: string;
  createdAt?: number;
  documents: DraftDocument[];
}

function record(value: unknown): value is Record<string, unknown> { return !!value && typeof value === 'object' && !Array.isArray(value); }
function bytes(value: unknown, label: string) {
  if (!Array.isArray(value) || value.length > MAX_DRAFT_SIZE) throw new Error(label + '数据无效');
  for (const part of value) if (!Number.isInteger(part) || (part as number) < 0 || (part as number) > 255) throw new Error(label + '数据无效');
  return Uint8Array.from(value as number[]);
}

/** 读取一份文档草稿的 Yjs 状态，校验结构并提取预览信息。 */
function describe(sourceId: string, update: Uint8Array, fallbackTitle?: string): DraftDocument {
  const doc = new Y.Doc();
  let store: BlockStore | undefined;
  try {
    try { Y.applyUpdate(doc, update); } catch { throw new Error('草稿包含损坏的文档数据'); }
    store = new BlockStore(doc);
    store.initialize(); store.ensureNote();
    const surface = doc.getMap<SurfaceElement>('surface');
    if (!validSurface(new Map(surface.entries()))) throw new Error('草稿白板数据无效');
    const node = readDocument(store);
    const missing = new Set(documentAssets(node));
    for (const element of surface.values()) if (element.kind === 'image') missing.add(element.assetId);
    const markdown = markdownSerializer.serialize(node).trim().replace(/\s+/g, ' ');
    return {
      sourceId,
      title: store.meta.get('title') || fallbackTitle || '无标题文档',
      summary: markdown.length > SUMMARY_LENGTH ? markdown.slice(0, SUMMARY_LENGTH) + '…' : markdown,
      surfaceElements: surface.size,
      missingAssets: [...missing],
      update,
    };
  } finally { store?.destroy(); doc.destroy(); }
}

/**
 * 解析草稿文件文本，返回可直接展示的恢复预览。
 * 同时接受工作空间草稿与单文档草稿；格式、版本、大小与 Yjs 数据都在此校验。
 */
export function parseDraft(text: string): DraftPreview {
  if (new Blob([text]).size > MAX_DRAFT_SIZE) throw new Error('草稿文件超过 50 MB，无法恢复');
  let data: unknown;
  try { data = JSON.parse(text); } catch { throw new Error('无法读取草稿 JSON 文件'); }
  if (!record(data) || data.version !== 1) throw new Error('不是支持的草稿文件，或草稿版本不受支持');
  if (data.format === 'mydoc-document-draft') {
    if (typeof data.id !== 'string' || !data.id) throw new Error('草稿文档标识无效');
    return { kind: 'document', documents: [describe(data.id, bytes(data.update, '草稿文档'))] };
  }
  if (data.format !== 'mydoc-workspace-draft') throw new Error('不是支持的草稿文件，或草稿版本不受支持');
  if (typeof data.workspaceId !== 'string' || !data.workspaceId) throw new Error('草稿工作空间标识无效');
  if (data.createdAt !== undefined && (typeof data.createdAt !== 'number' || !Number.isFinite(data.createdAt))) throw new Error('草稿创建时间无效');
  if (!Array.isArray(data.documents) || data.documents.length > MAX_DRAFT_DOCUMENTS) throw new Error('草稿文档列表无效');
  // 工作空间草稿的 documents 可能是仅 ID 的字符串数组（没有已打开文档时），此时没有可恢复内容。
  const documents: DraftDocument[] = [];
  const seen = new Set<string>();
  let updateBytes = 0;
  for (const entry of data.documents as unknown[]) {
    if (typeof entry === 'string') continue;
    if (!record(entry) || typeof entry.id !== 'string' || !entry.id) throw new Error('草稿文档信息无效');
    if (entry.title !== undefined && typeof entry.title !== 'string') throw new Error('草稿文档标题无效');
    if (seen.has(entry.id)) throw new Error('草稿包含重复的文档');
    seen.add(entry.id);
    const update = bytes(entry.update, '草稿文档');
    updateBytes += update.byteLength;
    if (updateBytes > MAX_DRAFT_SIZE) throw new Error('草稿中的文档数据超过 50 MB，无法恢复');
    documents.push(describe(entry.id, update, entry.title as string | undefined));
  }
  if (!documents.length) throw new Error('草稿中没有可恢复的文档内容；导出时可能没有打开任何文档');
  return { kind: 'workspace', sourceWorkspaceId: data.workspaceId, createdAt: data.createdAt as number | undefined, documents };
}

export interface DraftRestoreResult { id: string; title: string; missingAssets: string[] }

/**
 * 把选中的草稿文档恢复为目标工作区中的新文档。
 *
 * 一律创建新 ID 的新文档，不写回原文档、不复用原 ID，因此被永久删除或已失去权限的
 * 文档只会得到一份副本，不会绕过 tombstone 或重新取得访问权。目标必须是个人工作区。
 */
export async function restoreDraft(workspace: LocalWorkspace, documents: DraftDocument[], folderId: string | null = null): Promise<DraftRestoreResult[]> {
  if (workspace.sync) throw new Error('草稿只能恢复到个人工作区，请先切换工作区');
  workspace.assertWritable();
  if (!documents.length) throw new Error('请选择要恢复的文档');
  const results: DraftRestoreResult[] = [];
  for (const entry of documents) {
    // createSnapshot 生成新 ID 并按新工作区重建元信息；草稿没有原创建时间，使用恢复时间。
    const id = await workspace.createSnapshot(entry.title, Date.now(), folderId, entry.update);
    // 草稿不含附件字节：目标工作区已有同 ID 文件时引用仍然有效，否则作为缺失提示。
    const missing: string[] = [];
    for (const assetId of entry.missingAssets) if (!await workspace.assets.get(assetId)) missing.push(assetId);
    results.push({ id, title: entry.title, missingAssets: missing });
  }
  return results;
}
