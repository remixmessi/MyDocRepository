export interface AssetRecord { id: string; name: string; mime: string; size: number; bytes: Uint8Array }
export interface AssetStorage {
  get(workspace: string, id: string): Promise<AssetRecord | null>;
  put(workspace: string, asset: AssetRecord): Promise<void>;
  remove?(workspace:string,id:string):Promise<void>;
  removeWorkspace(workspace: string): Promise<void>;
  close(): void;
}
export class IndexedDBAssetStorage implements AssetStorage {
  private db: Promise<IDBDatabase>;
  constructor(name = 'mydoc-assets-v1') {
    this.db = new Promise((resolve, reject) => { const request = indexedDB.open(name, 1); request.onupgradeneeded = () => request.result.createObjectStore('assets'); request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); request.onblocked = () => reject(new Error('文件数据库被其他标签页阻塞，请关闭后重试')); });
  }
  async get(workspace: string, id: string) {
    const db = await this.db; return new Promise<AssetRecord | null>((resolve, reject) => { const tx = db.transaction('assets', 'readonly'); const request = tx.objectStore('assets').get([workspace, id]); tx.oncomplete = () => resolve(request.result ?? null); tx.onerror = tx.onabort = () => reject(tx.error); });
  }
  async put(workspace: string, asset: AssetRecord) {
    const db = await this.db; return new Promise<void>((resolve, reject) => { const tx = db.transaction('assets', 'readwrite'); tx.objectStore('assets').put(asset, [workspace, asset.id]); tx.oncomplete = () => resolve(); tx.onerror = tx.onabort = () => reject(tx.error ?? new Error('文件保存失败')); });
  }
  async remove(workspace:string,id:string) { const db=await this.db;return new Promise<void>((resolve,reject)=>{const tx=db.transaction('assets','readwrite');tx.objectStore('assets').delete([workspace,id]);tx.oncomplete=()=>resolve();tx.onerror=tx.onabort=()=>reject(tx.error);}); }
  async removeWorkspace(workspace: string) {
    const db = await this.db; return new Promise<void>((resolve, reject) => { const tx = db.transaction('assets', 'readwrite'); const table = tx.objectStore('assets'); const cursor = table.openCursor(); cursor.onsuccess = () => { const value = cursor.result; if (!value) return; if (Array.isArray(value.key) && value.key[0] === workspace) value.delete(); value.continue(); }; tx.oncomplete = () => resolve(); tx.onerror = tx.onabort = () => reject(tx.error); });
  }
  close() { void this.db.then(db => db.close(), () => {}); }
}
export const MAX_ASSET_SIZE = 20 * 1024 * 1024;
export function imageMime(bytes: Uint8Array): string | null {
  if (bytes.length >= 8 && [137,80,78,71,13,10,26,10].every((value, i) => bytes[i] === value)) return 'image/png';
  if (bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255) return 'image/jpeg';
  const header = String.fromCharCode(...bytes.subarray(0, 12));
  if (header.startsWith('GIF87a') || header.startsWith('GIF89a')) return 'image/gif';
  if (header.startsWith('RIFF') && header.slice(8) === 'WEBP') return 'image/webp'; return null;
}
export class WorkspaceAssets {
  private pending = new Set<Promise<AssetRecord>>(); private cleared = false;
  constructor(private storage: AssetStorage, readonly workspace: string) {}
  get saving() { return this.pending.size>0; }
  get(id: string) { return this.storage.get(this.workspace, id); }
  put(asset: AssetRecord) { return this.storage.put(this.workspace, asset); }
  remove(id:string) { return this.storage.remove?.(this.workspace,id)??Promise.resolve(); }
  async clear() { this.cleared = true; await Promise.allSettled([...this.pending]); await this.storage.removeWorkspace(this.workspace); }
  save(file: File, asImage: boolean) {
    if (this.cleared) return Promise.reject(new Error('工作区已删除，无法继续保存文件'));
    const operation = this.saveFile(file, asImage); this.pending.add(operation); void operation.then(() => this.pending.delete(operation), () => this.pending.delete(operation)); return operation;
  }
  private async saveFile(file: File, asImage: boolean) {
    if (file.size > MAX_ASSET_SIZE) throw new Error('单个文件不能超过 20 MB');
    const bytes = new Uint8Array(await file.arrayBuffer()); const detected = imageMime(bytes);
    if (asImage && !detected) throw new Error('图片支持 PNG、JPEG、GIF、WebP；其他格式可作为附件添加');
    const asset: AssetRecord = { id: crypto.randomUUID(), name: file.name || '未命名文件', mime: asImage ? detected! : file.type || 'application/octet-stream', size: bytes.length, bytes };
    await this.put(asset); return asset;
  }
}
