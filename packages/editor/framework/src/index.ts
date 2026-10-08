import { Y, LOCAL_ORIGIN } from '@mydoc/crdt';

export interface BlockSchema {
  flavour: string;
  role: 'root' | 'container' | 'content';
  parents?: string[];
  defaults?: Record<string, unknown>;
  validate: (props: Record<string, unknown>) => void;
}
export class BlockRegistry {
  private schemas = new Map<string, BlockSchema>();
  register(schema: BlockSchema) {
    if (this.schemas.has(schema.flavour)) throw new Error(`重复块类型：${schema.flavour}`);
    this.schemas.set(schema.flavour, schema);
    return this;
  }
  get(flavour: string) {
    const schema = this.schemas.get(flavour);
    if (!schema) throw new Error(`未知块类型：${flavour}`);
    return schema;
  }
}
const noProps = (props: Record<string, unknown>) => { if (Object.keys(props).length) throw new Error('此块无可配置属性'); };
export const createRegistry = () => new BlockRegistry()
  .register({ flavour: 'root', role: 'root', validate: noProps })
  .register({ flavour: 'note', role: 'container', parents: ['root'], validate: noProps })
  .register({ flavour: 'paragraph', role: 'content', parents: ['root', 'note'], validate: noProps })
  .register({ flavour: 'heading', role: 'content', parents: ['root', 'note'], defaults: { level: 1 }, validate: props => {
    if (!Number.isInteger(props.level) || (props.level as number) < 1 || (props.level as number) > 6 || Object.keys(props).some(key => key !== 'level')) throw new Error('标题级别无效');
  } })
  .register({ flavour: 'quote', role: 'content', parents: ['root', 'note'], validate: noProps })
  .register({ flavour: 'list', role: 'content', parents: ['root', 'note', 'list'], defaults: { type: 'bullet', checked: false }, validate: props => {
    if (!['bullet', 'ordered', 'todo'].includes(props.type as string) || typeof props.checked !== 'boolean' || Object.keys(props).some(key => !['type', 'checked'].includes(key))) throw new Error('列表属性无效');
  } })
  .register({ flavour: 'code', role: 'content', parents: ['root', 'note'], defaults: { language: 'javascript' }, validate: props => {
    if (!['javascript', 'typescript', 'json', 'css', 'markup', 'plain'].includes(props.language as string) || Object.keys(props).some(key => key !== 'language')) throw new Error('代码语言无效');
  } })
  .register({ flavour: 'divider', role: 'content', parents: ['root', 'note'], validate: noProps });

type Block = Y.Map<unknown>;
export class BlockStore {
  readonly blocks: Y.Map<Block>;
  readonly meta: Y.Map<string>;
  readonly history: Y.UndoManager;
  constructor(readonly doc: Y.Doc, readonly registry = createRegistry()) {
    this.blocks = doc.getMap<Block>('blocks');
    this.meta = doc.getMap<string>('meta');
    this.history = new Y.UndoManager([this.blocks, this.meta], { trackedOrigins: new Set([LOCAL_ORIGIN]), captureTimeout: 500 });
  }
  initialize() {
    if (this.blocks.size) return;
    this.doc.transact(() => {
      this.addBlock('root', null, 0, {}, 'root');
      this.addBlock('note', 'root', 0, {}, 'note-main');
      this.meta.set('title', '无标题文档');
    }, LOCAL_ORIGIN);
    this.history.clear();
  }
  block(id: string): Block {
    const block = this.blocks.get(id);
    if (!block) throw new Error(`块不存在：${id}`);
    return block;
  }
  children(id = 'root'): string[] {
    return (this.block(id).get('children') as Y.Array<string>).toArray();
  }
  fragment(id: string) { return this.block(id).get('content') as Y.XmlFragment; }
  parent(id: string): string | null {
    for (const [parentId] of this.blocks) if (this.children(parentId).includes(id)) return parentId;
    return null;
  }
  addBlock(flavour: string, parent: string | null, index?: number, props: Record<string, unknown> = {}, id: string = crypto.randomUUID()): string {
    const schema = this.registry.get(flavour);
    props = { ...schema.defaults, ...props };
    schema.validate(props);
    if (this.blocks.has(id)) throw new Error('块 ID 已存在');
    if (schema.role === 'root') {
      if (parent !== null || this.blocks.size) throw new Error('文档只能有一个根块');
    } else {
      if (parent === null) throw new Error('内容块必须有父块');
      this.validateParent(flavour, parent);
    }
    const siblings = parent === null ? null : this.block(parent).get('children') as Y.Array<string>;
    const position = index ?? siblings?.length ?? 0;
    if (!Number.isInteger(position) || position < 0 || position > (siblings?.length ?? 0)) throw new Error('无效块位置');
    this.history.stopCapturing();
    this.doc.transact(() => {
      const block = new Y.Map<unknown>();
      block.set('flavour', flavour);
      block.set('children', new Y.Array<string>());
      block.set('props', new Y.Map(Object.entries(props)));
      if (schema.role === 'content') block.set('content', new Y.XmlFragment());
      this.blocks.set(id, block);
      siblings?.insert(position, [id]);
    }, LOCAL_ORIGIN);
    this.history.stopCapturing();
    return id;
  }
  updateBlock(id: string, props: Record<string, unknown>) {
    const block = this.block(id);
    const values = block.get('props') as Y.Map<unknown>;
    this.registry.get(block.get('flavour') as string).validate({ ...values.toJSON(), ...props });
    this.history.stopCapturing();
    this.doc.transact(() => Object.entries(props).forEach(([key, value]) => values.set(key, value)), LOCAL_ORIGIN);
    this.history.stopCapturing();
  }
  deleteBlock(id: string) {
    const parent = this.parent(id);
    if (!parent) throw new Error('不能删除根块');
    this.history.stopCapturing();
    this.doc.transact(() => {
      const siblings = this.block(parent).get('children') as Y.Array<string>;
      siblings.delete(siblings.toArray().indexOf(id));
      const remove = (target: string) => { this.children(target).forEach(remove); this.blocks.delete(target); };
      remove(id);
    }, LOCAL_ORIGIN);
    this.history.stopCapturing();
  }
  // index 指删除原位置后，在目标父块中的插入位置。
  moveBlock(id: string, parent: string, index: number) {
    const source = this.parent(id);
    if (!source) throw new Error('不能移动根块');
    this.validateParent(this.flavour(id), parent);
    for (let ancestor: string | null = parent; ancestor; ancestor = this.parent(ancestor)) {
      if (ancestor === id) throw new Error('不能把块移入自身或后代');
    }
    const target = this.block(parent).get('children') as Y.Array<string>;
    const max = target.length - (source === parent ? 1 : 0);
    if (!Number.isInteger(index) || index < 0 || index > max) throw new Error('无效块位置');
    this.history.stopCapturing();
    this.doc.transact(() => {
      const siblings = this.block(source).get('children') as Y.Array<string>;
      siblings.delete(siblings.toArray().indexOf(id));
      target.insert(index, [id]);
    }, LOCAL_ORIGIN);
    this.history.stopCapturing();
  }
  flavour(id: string) { return this.block(id).get('flavour') as string; }
  props(id: string) { return (this.block(id).get('props') as Y.Map<unknown>).toJSON() as Record<string, unknown>; }
  private validateParent(flavour: string, parent: string) {
    const schema = this.registry.get(flavour);
    const parentFlavour = this.flavour(parent);
    if (!(schema.parents ?? ['root']).includes(parentFlavour)) throw new Error('块类型不允许此父块');
  }
  /** 多个块/富文本修改合并为一个撤销单元。 */
  transact(action: () => void) {
    this.history.stopCapturing();
    try { this.doc.transact(action, LOCAL_ORIGIN); } finally { this.history.stopCapturing(); }
  }
  convertBlock(id: string, flavour: string, props: Record<string, unknown> = {}) {
    const parent = this.parent(id);
    if (!parent || this.flavour(id) === 'note') throw new Error('不能转换结构块');
    this.validateParent(flavour, parent);
    if (this.children(id).length && flavour !== 'list') throw new Error('请先移出子列表再转换类型');
    const schema = this.registry.get(flavour);
    if (schema.role !== 'content') throw new Error('目标必须为内容块');
    const values = { ...schema.defaults, ...props }; schema.validate(values);
    this.transact(() => {
      this.block(id).set('flavour', flavour);
      this.block(id).set('props', new Y.Map(Object.entries(values)));
      // 分隔线保留原 fragment 以支持无损转换/撤销。
      if (!this.block(id).has('content')) this.block(id).set('content', new Y.XmlFragment());
    });
  }
  indent(id: string) {
    if (this.flavour(id) !== 'list') return false;
    const parent = this.parent(id)!; const siblings = this.children(parent); const index = siblings.indexOf(id);
    const previous = siblings[index - 1];
    if (!previous || this.flavour(previous) !== 'list') return false;
    this.moveBlock(id, previous, this.children(previous).length); return true;
  }
  outdent(id: string) {
    const parent = this.parent(id);
    if (!parent || this.flavour(parent) !== 'list') return false;
    const grandparent = this.parent(parent)!;
    this.moveBlock(id, grandparent, this.children(grandparent).indexOf(parent) + 1); return true;
  }
  contentIds(parent = 'root'): string[] {
    return this.children(parent).flatMap(id => this.flavour(id) === 'note' ? this.contentIds(id) : [id, ...this.contentIds(id)]);
  }
  /** 第一批粘贴产生的多个内部段落规范化为软换行，保留 XmlText 格式。 */
  normalizeText(id: string) {
    const fragment = this.fragment(id);
    if (!fragment || fragment.length < 2) return;
    this.doc.transact(() => {
      const paragraph = new Y.XmlElement('paragraph');
      const children: (Y.XmlElement | Y.XmlText)[] = [];
      fragment.toArray().forEach((node, index) => {
        if (index) children.push(new Y.XmlElement('hard_break'));
        if (node instanceof Y.XmlElement) node.toArray().forEach(child => { if (child instanceof Y.XmlElement || child instanceof Y.XmlText) children.push(child.clone()); });
        else if (node instanceof Y.XmlText) children.push(node.clone());
      });
      paragraph.insert(0, children); fragment.delete(0, fragment.length); fragment.insert(0, [paragraph]);
    }, 'schema-migration-v2');
  }
  /** 升级第一批平铺文档；不改变块 ID、富文本或用户历史。 */
  ensureNote() {
    const note = this.children().find(id => this.flavour(id) === 'note');
    if (note) return note;
    let result = '';
    this.doc.transact(() => {
      const existing = this.children(); result = this.addBlock('note', 'root');
      existing.forEach(id => this.moveBlock(id, result, this.children(result).length));
    }, 'schema-migration-v2');
    return result;
  }
  rename(title: string) { this.doc.transact(() => this.meta.set('title', title), LOCAL_ORIGIN); }
  destroy() { this.history.destroy(); }
}
