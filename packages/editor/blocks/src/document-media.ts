import { Plugin, TextSelection, type SelectionBookmark } from 'prosemirror-state';
import type { EditorView, NodeView } from 'prosemirror-view';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import type { BlockStore } from '@mydoc/editor-framework';
import { imageMime, type AssetRecord, type WorkspaceAssets } from '@mydoc/nbstore';
import { documentSchema } from './document-format';
import {attachmentPlayer,playbackType,type AttachmentPlayer} from './media-playback';
import { isPdfAttachment } from './pdf-preview-utils';
export type UploadKind = 'image' | 'attachment' | 'auto';
export function formatFileSize(size: number) { return size < 1024 ? `${size} B` : size < 1024 * 1024 ? `${(size / 1024).toFixed(1)} KB` : `${(size / (1024 * 1024)).toFixed(1)} MB`; }
export function documentMedia(assets: WorkspaceAssets, store: BlockStore, notify: (text: string) => void, progress: (text: string) => void, previewPdf?: (id:string)=>void) {
  const players=new Set<AttachmentPlayer>();
  let editor: EditorView; let disposed = false; let uploading = false; let bookmark: SelectionBookmark | null = null;
  async function upload(files: File[], kind: UploadKind) {
    if (uploading || !files.length) return;
    if (files.length > 10) { notify('每次最多添加 10 个文件'); return; }
    if (editor.composing || editor.state.selection.$from.parent.type.spec.code) { notify('请在正文中选择文件插入位置'); return; }
    uploading = true; bookmark = editor.state.selection.getBookmark(); notify(''); const nodes: PMNode[] = []; const failures: string[] = [];
    try {
      for (const [index, file] of files.entries()) {
        if (disposed) break; progress(`正在保存文件 ${index + 1}/${files.length}…`);
        try { const asImage = kind === 'image' || (kind === 'auto' && /^(image\/(png|jpeg|gif|webp))$/i.test(file.type)); const asset = await assets.save(file, asImage);
          if (nodes.length) nodes.push(documentSchema.nodes.hard_break.create());
          nodes.push(asImage ? documentSchema.nodes.image.create({ src: `asset:${asset.id}`, alt: asset.name, title: asset.name }) : documentSchema.nodes.attachment.create({ assetId: asset.id, name: asset.name, size: asset.size, mime: asset.mime }));
        } catch (error) { failures.push(`${file.name}：${error instanceof Error ? error.message : '文件保存失败，请重试'}`); }
      }
      if (!disposed && nodes.length && bookmark) { const selection = bookmark.resolve(editor.state.doc); bookmark = null; store.history.stopCapturing(); editor.dispatch(editor.state.tr.setSelection(selection).replaceSelection(new Slice(Fragment.fromArray(nodes), 0, 0)).scrollIntoView()); store.history.stopCapturing(); editor.focus(); }
    } finally { uploading = false; bookmark = null; if (!disposed) { progress(''); if (failures.length) notify(failures.join('；')); } }
  }
  const plugin = new Plugin({ props: {
    handlePaste(_view, event) { const files = [...(event.clipboardData?.files ?? [])]; if (!files.length) return false; void upload(files, 'auto'); return true; },
    handleDrop(view, event) { const files = [...(event.dataTransfer?.files ?? [])]; if (!files.length) return false; event.preventDefault(); const position = view.posAtCoords({ left: event.clientX, top: event.clientY }); if (position) { view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(position.pos)))); } void upload(files, 'auto'); return true; },
  }, view(view) { editor = view; return { update(view) { editor = view;if(view.dom.closest('[hidden]'))for(const player of players)player.pause(); }, destroy() { disposed = true;for(const player of players)player.pause(); } }; }, state: { init: () => null, apply(tr) { if (bookmark) bookmark = bookmark.map(tr.mapping); return null; } } });
  function nodeView(node: PMNode): NodeView {
    const image = node.type.name === 'image'; const local = image && String(node.attrs.src).startsWith('asset:');
    const dom = document.createElement('span'); dom.contentEditable = 'false'; dom.className = image ? 'media-image' : 'media-attachment'; let url: string | null = null; let removed = false; let generation = 0; let loaded: AssetRecord | null = null;let player:AttachmentPlayer|null=null;
    const id = () => image ? String(node.attrs.src).slice(6) : node.attrs.assetId;
    const revoke = () => {if(player){players.delete(player);player.destroy();player=null;} if (url) URL.revokeObjectURL(url); url = null; };
    const download = () => { if (!loaded) return; const link = document.createElement('a'); const temporary = URL.createObjectURL(new Blob([loaded.bytes], { type: loaded.mime })); link.href = temporary; link.download = loaded.name; link.click(); setTimeout(() => URL.revokeObjectURL(temporary), 1000); };
    async function render() {
      const current = ++generation; revoke(); loaded = null; dom.replaceChildren();
      const label = document.createElement('span'); label.className = 'media-file-name'; label.textContent = image ? node.attrs.alt || '图片' : node.attrs.name; dom.append(label);
      if (image && !local) { const img = document.createElement('img'); img.src = node.attrs.src; img.alt = node.attrs.alt || ''; img.title = node.attrs.title || ''; dom.replaceChildren(img); return; }
      const status = document.createElement('span'); status.className = 'media-status'; status.role = 'status'; status.textContent = '正在读取文件…'; dom.append(status);
      const preview=()=>{if(image||!previewPdf)return;const button=document.createElement('button');button.type='button';button.textContent='预览 PDF';button.setAttribute('aria-label',`预览 PDF ${loaded?.name||node.attrs.name}`);button.addEventListener('click',()=>previewPdf(id()));dom.append(button);};
      try {
        const record = await assets.get(id()); if (removed || current !== generation) return; if (!record) throw new Error('本地文件缺失，请恢复包含文件的工作区备份'); if (image && !imageMime(record.bytes)) throw new Error('图片数据无法读取'); loaded = record;
        label.textContent = record.name; status.textContent = formatFileSize(record.size);
        if(!image){const type=playbackType(record.name,record.mime,record.bytes);if(type){player=attachmentPlayer(record,type,()=>{void render();});players.add(player);dom.classList.add('media-playable');dom.append(player.dom);}}
        if(isPdfAttachment(record.name,record.mime,record.bytes))preview();
        if (image) { const img = document.createElement('img'); url = URL.createObjectURL(new Blob([record.bytes], { type: imageMime(record.bytes)! })); img.src = url; img.alt = node.attrs.alt || record.name; img.addEventListener('error', () => { status.textContent = '图片无法显示，可下载原文件检查'; }); dom.prepend(img); }
        const button = document.createElement('button'); button.type = 'button'; button.textContent = '下载'; button.setAttribute('aria-label', `下载 ${record.name}`); button.addEventListener('click', download);if(player)dom.insertBefore(button,player.dom);else dom.append(button);
      } catch (error) { if (removed || current !== generation) return; status.textContent = error instanceof Error ? error.message : '读取文件失败'; dom.classList.add('media-error');if(isPdfAttachment(node.attrs.name,node.attrs.mime))preview(); const retry = document.createElement('button'); retry.type = 'button'; retry.textContent = '重试'; retry.addEventListener('click', () => { dom.classList.remove('media-error'); void render(); }); dom.append(retry); }
    }
    void render(); return { dom, stopEvent: event => !!(event.target as HTMLElement)?.closest('button,select,input,audio,video,.attachment-player'), ignoreMutation: () => true, destroy() { removed = true; generation++; revoke(); } };
  }
  return { plugin, upload, nodeView };
}
