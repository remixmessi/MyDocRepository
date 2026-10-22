import {cloudRequest} from '../../../common/nbstore/src/cloud';
import {isCloudWorkspaceArchive,type CloudWorkspaceArchive} from '../../../common/nbstore/src/workspace-archive';
import {validWebClip} from './clip-source';
import {validJournalDate} from './journal';
import {validCommentThread,MAX_COMMENT_THREADS,type CommentThread} from '../../../editor/blocks/src/comment-model';
import {validFootnoteLabel} from '../../../editor/blocks/src/footnote-format';
import { validImportSource } from './document-import';
import { validDocumentTemplate, templateStore, disposeTemplate, snapshotAssets, MAX_TEMPLATES, type DocumentTemplate } from './document-templates';
import { validTag, validCollection, uniqueIds, type TagMeta, type CollectionMeta } from './organization';
import{validExpression}from'../../../editor/blocks/src/expression-format';
import {validTemplate,templateAssets,type BoardTemplate} from '@mydoc/editor-blocks';
import { validSurface,type SurfaceElement } from '@mydoc/editor-blocks';
import { MAX_ASSET_SIZE,imageMime } from '@mydoc/nbstore';
import { Y } from '@mydoc/crdt';
import { BlockStore } from '@mydoc/editor-framework';
import { validWebCard, validDocEmbed, validSurfaceRef } from '@mydoc/editor-blocks';
import { readDocument, documentAssets } from '@mydoc/editor-blocks';
import type { LocalWorkspace, FolderMeta, DocumentMeta } from './workspace';
export interface WorkspaceBackup { type: 'mydoc-workspace'; version: 1|2; office?:CloudWorkspaceArchive; officeOrganization?:{id:string;folderId?:string|null;favorite?:boolean;tagIds?:string[]}[]; workspaceId?: string; name: string; exportedAt: number; folders: FolderMeta[]; templates?: DocumentTemplate[]; tags?: TagMeta[]; collections?: CollectionMeta[]; assets?: { id: string; name: string; mime: string; size: number; data: string }[]; documents: { meta: DocumentMeta; update: string }[]; trash?: { meta: DocumentMeta; update: string; deletedAt: number; collectionIds: string[] }[] }
export function encodeBinary(bytes: Uint8Array) { let value = ''; for (let index = 0; index < bytes.length; index += 8192) value += String.fromCharCode(...bytes.subarray(index, index + 8192)); return btoa(value); }
export function decodeBinary(value: string) { if (!value || !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) throw new Error('备份包含无效的文档数据'); const text = atob(value); return Uint8Array.from(text, char => char.charCodeAt(0)); }
export async function exportBackup(workspace: LocalWorkspace, name: string): Promise<WorkspaceBackup> {
  await workspace.prepareWorkspaceBackup();
  const office=!workspace.sync&&workspace.officeAvailable?await (await cloudRequest('/office-workspace-backup?'+new URLSearchParams({workspace:workspace.id}))).json() as CloudWorkspaceArchive:undefined;
  if(workspace.officeResources.size&&!office)throw new Error('当前存储无法生成办公文件完整备份');
  await workspace.flush(); const documents: WorkspaceBackup['documents'] = []; const trash: NonNullable<WorkspaceBackup['trash']> = []; const assetIds = new Set<string>(); const assets: NonNullable<WorkspaceBackup['assets']> = [];
  const assetOwners=new Map<string,string>();
  for (const meta of workspace.allDocuments()) { const item = await workspace.openDocument(meta.id); documentAssets(readDocument(item.store)).forEach(id => {assetIds.add(id);assetOwners.set(id,meta.id);}); for(const image of item.store.doc.getMap<SurfaceElement>('surface').values()) if(image.kind==='image'){assetIds.add(image.assetId);assetOwners.set(image.assetId,meta.id);} for(const template of item.store.doc.getMap<BoardTemplate>('whiteboardTemplates').values()){if(!validTemplate(template))throw new Error('白板模板无效');templateAssets(template).forEach(id=>{assetIds.add(id);assetOwners.set(id,meta.id);});} const update=encodeBinary(Y.encodeStateAsUpdate(item.store.doc));const recycled=workspace.trash.get(meta.id);if(recycled)trash.push({...recycled,update});else documents.push({meta,update}); }
  for (const id of assetIds) { const asset = await (assetOwners.has(id)?workspace.documentAssets(assetOwners.get(id)!).get(id):workspace.assets.get(id)); if (!asset) throw new Error('本地文件缺失，无法生成完整备份，请先恢复文件'); assets.push({ id, name: asset.name, mime: asset.mime, size: asset.size, data: encodeBinary(asset.bytes) }); }
  for(const template of workspace.templates.values()){const store=templateStore(template);try{for(const id of snapshotAssets(store))if(!assetIds.has(id)){const asset=await workspace.assets.get(id);if(!asset)throw new Error('模板文件缺失，无法生成完整备份');assetIds.add(id);assets.push({id,name:asset.name,mime:asset.mime,size:asset.size,data:encodeBinary(asset.bytes)});}}finally{disposeTemplate(store);}}
  const result: WorkspaceBackup = { type: 'mydoc-workspace', version: office?2:1, ...(office?{office,officeOrganization:[...workspace.officeOrganization].filter(([id])=>office.office.some(r=>r.id===id&&!r.deleted)).map(([id,value])=>({id,...value}))}:{}), assets, workspaceId: workspace.id, name, exportedAt: Date.now(), templates:[...workspace.templates.values()], folders: workspace.folderList(), tags: workspace.organization.tagList(), collections: workspace.organization.collectionList(), documents, trash };
  if (new Blob([JSON.stringify(result)]).size > 50 * 1024 * 1024) throw new Error('完整备份超过 50 MB，请将文档分到不同工作区后备份'); return result;
}
function record(value: unknown): value is Record<string, unknown> { return !!value && typeof value === 'object' && !Array.isArray(value); }
export function parseBackup(text: string): WorkspaceBackup {
  if(new Blob([text]).size>50*1024*1024)throw new Error('完整备份超过 50 MB');
  let data: unknown; try { data = JSON.parse(text); } catch { throw new Error('无法读取备份 JSON 文件'); }
  if (!record(data) || (data.workspaceId !== undefined && typeof data.workspaceId !== 'string') || data.type !== 'mydoc-workspace' || ![1,2].includes(Number(data.version)) || typeof data.name !== 'string' || !data.name.trim() || !Array.isArray(data.documents) || !Array.isArray(data.folders) || data.documents.length > 10000 || data.folders.length > 10000) throw new Error('不是支持的文档库备份文件');
  if(data.trash!==undefined&&(!Array.isArray(data.trash)||data.trash.length+data.documents.length>10000))throw new Error('备份回收站无效');
  for(const entry of (data.trash??[]) as unknown[])if(!record(entry)||typeof entry.deletedAt!=='number'||!Number.isFinite(entry.deletedAt)||entry.deletedAt<=0||!uniqueIds(entry.collectionIds))throw new Error('备份回收站状态无效');
    if(data.templates!==undefined&&(!Array.isArray(data.templates)||data.templates.length>MAX_TEMPLATES))throw new Error('备份文档模板无效');
  const templates:DocumentTemplate[]=[],templateIds=new Set<string>(),templateNames=new Set<string>();
  for(const t of (data.templates??[]) as unknown[]){if(!validDocumentTemplate(t)||templateIds.has(t.id)||templateNames.has(t.name.trim().toLocaleLowerCase()))throw new Error('备份文档模板无效');templateIds.add(t.id);templateNames.add(t.name.trim().toLocaleLowerCase());templates.push(t);}
  const assetIds = new Set<string>(),imageAssetIds=new Set<string>();
  if (data.assets !== undefined && (!Array.isArray(data.assets) || data.assets.length > 10000)) throw new Error('备份文件列表无效');
  for (const asset of (data.assets ?? []) as unknown[]) {
    if (!record(asset) || typeof asset.id !== 'string' || !asset.id || assetIds.has(asset.id) || typeof asset.name !== 'string' || !asset.name || typeof asset.mime !== 'string' || typeof asset.size !== 'number' || !Number.isInteger(asset.size) || asset.size < 0 || asset.size > MAX_ASSET_SIZE || typeof asset.data !== 'string') throw new Error('备份文件信息无效');
    const bytes = asset.data ? decodeBinary(asset.data) : new Uint8Array(); if (bytes.length !== asset.size) throw new Error('备份文件大小与数据不一致'); assetIds.add(asset.id); if(imageMime(bytes))imageAssetIds.add(asset.id);
  }
  const tagIds = new Set<string>(); const tagNames = new Set<string>();
  if (data.tags !== undefined && (!Array.isArray(data.tags) || data.tags.length > 10000)) throw new Error('备份标签无效');
  for (const tag of (data.tags ?? []) as unknown[]) { if (!validTag(tag) || tagIds.has(tag.id) || tagNames.has(tag.name.trim().toLocaleLowerCase())) throw new Error('备份标签无效'); tagIds.add(tag.id); tagNames.add(tag.name.trim().toLocaleLowerCase()); }
  const folders = new Map<string, FolderMeta>(); const ids = new Set<string>(); const journalDates=new Set<string>();
  for (const value of data.folders) {
    if (!record(value) || typeof value.id !== 'string' || !value.id || folders.has(value.id) || typeof value.name !== 'string' || !value.name.trim() || (value.parentId !== null && typeof value.parentId !== 'string') || typeof value.createdAt !== 'number' || !Number.isFinite(value.createdAt)) throw new Error('备份文件夹信息无效');
    folders.set(value.id, value as unknown as FolderMeta);
  }
  for (const folder of folders.values()) { const visited = new Set<string>([folder.id]); let parent = folder.parentId; while (parent) { if (visited.has(parent) || !folders.has(parent)) throw new Error('备份文件夹存在循环或缺失的上级'); visited.add(parent); parent = folders.get(parent)!.parentId; } }
  for (const {entry,isTemplate} of [...[...data.documents,...(data.trash??[]) as unknown[]].map(entry=>({entry,isTemplate:false})),...templates.map(t=>({entry:{meta:{id:t.sourceDocId,title:t.name,createdAt:t.createdAt},update:t.update},isTemplate:true}))]) {
    if (!record(entry) || !record(entry.meta) || typeof entry.meta.id !== 'string' || !entry.meta.id || (!isTemplate&&ids.has(entry.meta.id)) || typeof entry.meta.title !== 'string' || typeof entry.meta.createdAt !== 'number' || !Number.isFinite(entry.meta.createdAt) || (entry.meta.folderId != null && (typeof entry.meta.folderId !== 'string' || !folders.has(entry.meta.folderId))) || typeof entry.update !== 'string') throw new Error('备份文档信息无效');
    if(entry.meta.journalDate!==undefined){if(!validJournalDate(entry.meta.journalDate)||journalDates.has(entry.meta.journalDate))throw new Error('备份日记日期无效或重复');journalDates.add(entry.meta.journalDate);}
    if(entry.meta.webClip!==undefined&&!validWebClip(entry.meta.webClip))throw new Error('备份剪藏来源无效');
    if(entry.meta.importSource!==undefined&&!validImportSource(entry.meta.importSource))throw new Error('备份导入来源无效');
    if ((entry.meta.favorite !== undefined && typeof entry.meta.favorite !== 'boolean') || (entry.meta.tagIds !== undefined && (!uniqueIds(entry.meta.tagIds) || entry.meta.tagIds.some(id => !tagIds.has(id))))) throw new Error('备份文档标签或收藏无效');
    if(!isTemplate)ids.add(entry.meta.id); const doc = new Y.Doc(); let store: BlockStore | undefined;
    try { Y.applyUpdate(doc, decodeBinary(entry.update)); store = new BlockStore(doc); if (!store.blocks.has('root') || store.flavour('root') !== 'root') throw new Error('缺失文档结构'); store.ensureNote();
      const comments=doc.getMap<CommentThread>('documentComments');if(comments.size>MAX_COMMENT_THREADS||[...comments].some(([id,thread])=>!validCommentThread(thread)||id!==thread.id))throw new Error('正文评论无效');
      const surface = doc.getMap<SurfaceElement>('surface'); if (!validSurface(new Map(surface.entries()))) throw new Error('白板数据无效'); for(const image of surface.values()){if(image.kind==='image'&&!imageAssetIds.has(image.assetId))throw new Error('引用的本地图片缺失');}
      const templates=doc.getMap<BoardTemplate>('whiteboardTemplates');if(templates.size>20)throw new Error('模板数量超限');for(const [id,template] of templates){if(!validTemplate(template)||id!==template.id)throw new Error('白板模板无效');for(const id of templateAssets(template))if(!imageAssetIds.has(id))throw new Error('模板图片缺失');}const document = readDocument(store); for (const id of documentAssets(document)) if (!assetIds.has(id)) throw new Error('引用的本地文件缺失'); document.descendants(node => {
        if(node.type.name.startsWith('footnote_')&&!validFootnoteLabel(node.attrs.label))throw new Error('脚注信息无效');
        if(['callout','math_inline','math_block','mermaid'].includes(node.type.name)&&!validExpression(node.type.name,node.attrs))throw new Error('表达内容无效');
        if(node.type.name==='surface_ref'&&!validSurfaceRef(node.attrs))throw new Error('画布引用无效');
        if (node.type.name === 'doc_embed' && !validDocEmbed(node.attrs)) throw new Error('内部文档嵌入无效');
        if (node.type.name === 'web_card' && !validWebCard(node.attrs)) throw new Error('网页卡片信息无效');
        if (node.type.name === 'attachment' && (typeof node.attrs.assetId !== 'string' || !node.attrs.assetId || typeof node.attrs.name !== 'string' || typeof node.attrs.size !== 'number' || node.attrs.size < 0)) throw new Error('附件信息无效');
        if (node.type.name === 'doc_reference' && (typeof node.attrs.docId !== 'string' || !node.attrs.docId || typeof node.attrs.workspaceId !== 'string' || !node.attrs.workspaceId || typeof node.attrs.label !== 'string')) throw new Error('页面引用无效');
        if (node.type.name !== 'code_block') return;
        const attrs = node.attrs;
        if ((attrs.id !== null && typeof attrs.id !== 'string') || typeof attrs.wrap !== 'boolean' || typeof attrs.lineNumbers !== 'boolean' || typeof attrs.collapsed !== 'boolean' || !Array.isArray(attrs.comments) || attrs.comments.some((comment: unknown) => !record(comment) || typeof comment.id !== 'string' || typeof comment.text !== 'string' || typeof comment.createdAt !== 'number' || !Number.isFinite(comment.createdAt))) throw new Error('代码块元数据无效');
      }); }
    catch { throw new Error('备份包含损坏的文档数据'); } finally { store?.destroy(); doc.destroy(); }
  }
  if(data.version===2){
    if(!isCloudWorkspaceArchive(data.office)||!Array.isArray(data.officeOrganization)||data.officeOrganization.length>10000)throw new Error('个人办公备份信息无效');
    const officeIds=new Set<string>();
    for(const r of data.office.office){if(!record(r)||typeof r.id!=='string'||!r.id.startsWith('doc:')||officeIds.has(r.id)||ids.has(r.id)||![0,1].includes(Number(r.deleted)))throw new Error('办公文件标识无效或重复');officeIds.add(r.id);if(!r.deleted)ids.add(r.id);}
    const seen=new Set<string>();for(const entry of data.officeOrganization){if(!record(entry)||typeof entry.id!=='string'||!officeIds.has(entry.id)||seen.has(entry.id)||Object.keys(entry).some(key=>!['id','folderId','favorite','tagIds'].includes(key))||entry.folderId!=null&&!folders.has(String(entry.folderId))||entry.favorite!==undefined&&typeof entry.favorite!=='boolean'||entry.tagIds!==undefined&&(!uniqueIds(entry.tagIds)||entry.tagIds.some(id=>!tagIds.has(id))))throw new Error('办公文件目录信息无效');seen.add(entry.id);}
  }else if(data.office!==undefined||data.officeOrganization!==undefined)throw new Error('旧版备份不能包含办公文件');
  const collectionIds = new Set<string>(); const collectionNames = new Set<string>();
  if (data.collections !== undefined && (!Array.isArray(data.collections) || data.collections.length > 10000)) throw new Error('备份集合无效');
  for (const value of (data.collections ?? []) as unknown[]) { if (!validCollection(value) || collectionIds.has(value.id) || collectionNames.has(value.name.trim().toLocaleLowerCase()) || value.documentIds.some(id => !ids.has(id)) || value.rules.tagIds.some(id => !tagIds.has(id))) throw new Error('备份集合或规则无效'); collectionIds.add(value.id); collectionNames.add(value.name.trim().toLocaleLowerCase()); }
  for(const entry of (data.trash??[]) as {collectionIds:string[]}[])if(entry.collectionIds.some(id=>!collectionIds.has(id)))throw new Error('备份回收站集合引用无效');
  return data as unknown as WorkspaceBackup;
}
export function downloadFile(name: string, text: string, type: string) { const url = URL.createObjectURL(new Blob([text], { type })); const link = document.createElement('a'); link.href = url; link.download = name; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
