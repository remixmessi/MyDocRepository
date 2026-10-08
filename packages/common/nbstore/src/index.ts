import { Y } from '@mydoc/crdt';

/** 存储层只处理 Yjs 二进制，平台实现不依赖 React 或编辑器。 */
export interface DocumentStorage {
  load(id: string): Promise<Uint8Array | null>;
  append(id: string, update: Uint8Array): Promise<void>;
  remove(id: string): Promise<void>;
  close(): void;
}
export class IndexedDBStorage implements DocumentStorage {
  private db: Promise<IDBDatabase>;
  constructor(name = 'mydoc-v1') {
    this.db = new Promise((resolve, reject) => {
      const request = indexedDB.open(name, 1);
      request.onupgradeneeded = () => request.result.createObjectStore('documents');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(new Error('数据库升级被其他标签页阻塞，请关闭其他文档库标签页后重试'));
    });
  }
  async load(id: string) {
    const db = await this.db;
    return new Promise<Uint8Array | null>((resolve, reject) => {
      const tx = db.transaction('documents', 'readonly');
      const request = tx.objectStore('documents').get(id);
      tx.oncomplete = () => resolve(request.result ?? null);
      tx.onerror = tx.onabort = () => reject(tx.error);
    });
  }
  async append(id: string, update: Uint8Array) {
    const db = await this.db;
    return new Promise<void>((resolve, reject) => {
      // 在同一 readwrite 事务中合并，防止多个实例覆盖彼此的 update。
      const tx = db.transaction('documents', 'readwrite');
      const table = tx.objectStore('documents');
      const request = table.get(id);
      request.onsuccess = () => {
        try { table.put(request.result ? Y.mergeUpdates([request.result, update]) : update, id); }
        catch { tx.abort(); }
      };
      tx.oncomplete = () => resolve();
      tx.onerror = tx.onabort = () => reject(tx.error ?? new Error('本地存储写入失败'));
    });
  }
  async remove(id: string) {
    const db = await this.db;
    return new Promise<void>((resolve, reject) => {
      const tx = db.transaction('documents', 'readwrite');
      tx.objectStore('documents').delete(id);
      tx.oncomplete = () => resolve();
      tx.onerror = tx.onabort = () => reject(tx.error);
    });
  }
  close() { void this.db.then(db => db.close(), () => {}); }
}
export type PersistenceState = 'loading' | 'saving' | 'saved' | 'error';
export class PersistentDocument {
  state: PersistenceState = 'loading';
  error: unknown;
  private queue: Promise<void> = Promise.resolve();
  private listeners = new Set<() => void>();
  private pending = 0;
  private disposed = false;
  constructor(readonly id: string, readonly doc: Y.Doc, private storage: DocumentStorage) {}
  subscribe = (listener: () => void) => { this.listeners.add(listener); return () => { this.listeners.delete(listener); }; };
  private notify() { this.listeners.forEach(listener => listener()); }
  async open(options: { load?: boolean } = {}) {
    try {
      const update = options.load === false ? null : await this.storage.load(this.id);
      if (this.disposed) return;
      if (update) Y.applyUpdate(this.doc, update);
      this.doc.on('update', this.onUpdate);
      this.state = 'saved';
      this.notify();
    } catch (error) { this.error = error; this.state = 'error'; this.notify(); throw error; }
  }
  private onUpdate = (update: Uint8Array) => {
    this.pending++;
    this.state = 'saving';
    this.notify();
    this.queue = this.queue.then(async () => {
      try {
        await this.storage.append(this.id, update);
      } catch (error) { this.error = error; }
      finally {
        this.pending--;
        this.state = this.error ? 'error' : this.pending ? 'saving' : 'saved';
        this.notify();
      }
    });
  };
  async retry() {
    await this.queue;
    this.error = undefined;
    this.onUpdate(Y.encodeStateAsUpdate(this.doc));
    await this.flush();
  }
  async flush() { await this.queue; if (this.error) throw this.error; }
  async destroy() {
    this.disposed = true;
    this.doc.off('update', this.onUpdate);
    await this.queue;
    this.listeners.clear();
    this.doc.destroy();
  }
}
export { IndexedDBAssetStorage, WorkspaceAssets, MAX_ASSET_SIZE, imageMime } from './assets';
export type { AssetStorage, AssetRecord } from './assets';
export {SQLiteStorage,SQLiteAssetStorage,projectPreferences,initializeProjectStorage} from './sqlite';

export type {DesktopBridge} from './desktop';
