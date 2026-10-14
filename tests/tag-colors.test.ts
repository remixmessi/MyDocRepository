import 'fake-indexeddb/auto';
import { expect, it } from 'vitest';
import { IndexedDBStorage } from '@mydoc/nbstore';
import { LocalWorkspace } from '../packages/frontend/core/src/workspace';
import { WorkspaceRegistry } from '../packages/frontend/core/src/workspaces';
import { emptyRules, type TagColor } from '../packages/frontend/core/src/organization';
import { exportBackup, parseBackup } from '../packages/frontend/core/src/backup';

it('旧标签可选色、单独改名保留颜色、恢复默认色，颜色不改变文档绑定和规则结果', async () => {
  const storage = crypto.randomUUID();
  let workspace = new LocalWorkspace(new IndexedDBStorage(storage));
  await workspace.open();
  try {
    const documentId = await workspace.create('颜色验证');
    const tagId = workspace.organization.saveTag('旧标签');
    const before = workspace.organization.tags.get(tagId)!;
    expect(before.color).toBeUndefined();
    workspace.organization.setDocumentTags(documentId, [tagId]);
    const collectionId = workspace.organization.saveCollection({ name: '规则', mode: 'rules', documentIds: [], rules: { ...emptyRules(), tagIds: [tagId] } });
    workspace.organization.saveTag('旧标签', tagId, 'blue');
    workspace.organization.saveTag('改过的标签', tagId);
    expect(workspace.organization.tags.get(tagId)).toEqual({ ...before, name: '改过的标签', color: 'blue' });
    for (const color of ['unknown', '#fff', '', 12, null]) {
      expect(() => workspace.organization.saveTag('错误颜色', undefined, color as TagColor)).toThrow('颜色无效');
    }
    expect(workspace.organization.tags.size).toBe(1);
    expect(workspace.documents.get(documentId)?.tagIds).toEqual([tagId]);
    expect(workspace.organization.members(collectionId).map(doc => doc.id)).toEqual([documentId]);
    await workspace.destroy();
    workspace = new LocalWorkspace(new IndexedDBStorage(storage));
    await workspace.open();
    expect(workspace.organization.tags.get(tagId)?.color).toBe('blue');
    workspace.organization.saveTag('改过的标签', tagId, 'default');
    expect(workspace.organization.tags.get(tagId)?.color).toBe('default');
  } finally { await workspace.destroy(); }
});

it('工作区备份重映射后保留颜色，兼容无颜色旧备份，拒绝损坏颜色字段', async () => {
  const registry = new WorkspaceRegistry(new IndexedDBStorage(crypto.randomUUID()));
  await registry.open();
  try {
    const workspace = await registry.get('local');
    const documentId = await workspace.create('备份文档', '保留正文');
    const tagId = workspace.organization.saveTag('项目', undefined, 'purple');
    workspace.organization.setDocumentTags(documentId, [tagId]);
    const backup = await exportBackup(workspace, '颜色备份');
    const restored = await registry.restore(parseBackup(JSON.stringify(backup)));
    const restoredTagId = restored.list()[0].tagIds![0];
    expect(restoredTagId).not.toBe(tagId);
    expect(restored.organization.tags.get(restoredTagId)).toMatchObject({ name: '项目', color: 'purple' });
    for (const color of ['unknown', '#fff', '', 12, null, {}]) {
      expect(() => parseBackup(JSON.stringify({ ...backup, tags: [{ ...backup.tags![0], color }] }))).toThrow('备份标签无效');
    }
    const legacy = { ...backup, tags: backup.tags!.map(({ color, ...tag }) => tag) };
    const restoredLegacy = await registry.restore(parseBackup(JSON.stringify(legacy)));
    expect(restoredLegacy.organization.tagList()[0].color).toBeUndefined();
    expect(restoredLegacy.list()[0].tagIds).toEqual([restoredLegacy.organization.tagList()[0].id]);
  } finally { await registry.destroy(); }
});
