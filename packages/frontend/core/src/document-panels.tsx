import { useState } from 'react';
import type { OutlineEntry } from '@mydoc/editor-blocks';
import type { LocalWorkspace } from './workspace';
export function DocumentOutline({ entries, onJump, onClose }: { entries: OutlineEntry[]; onJump: (position: number) => void; onClose: () => void }) {
  return <aside className="document-outline" aria-label="文档大纲"><div className="outline-heading"><h2>大纲</h2><button aria-label="关闭大纲" onClick={onClose}>×</button></div><nav aria-label="标题导航">{entries.length ? entries.map(entry => <button key={entry.position} title={entry.text} style={{ paddingLeft: 8 + (entry.level - 1) * 12 }} onClick={() => onJump(entry.position)}>{entry.text}</button>) : <p>使用标题后，这里会生成大纲。</p>}</nav></aside>;
}
export function Backlinks({ workspace, documentId, onNavigate }: { workspace: LocalWorkspace; documentId: string; onNavigate: (id: string, workspaceId: string) => void }) {
  const [open, setOpen] = useState(false);
  const incoming = workspace.references.backlinks(documentId, workspace.id).filter(link => workspace.documents.has(link.sourceId));
  const outgoing = [...new Map(workspace.references.outgoing(documentId).map(link => [`${link.workspaceId}:${link.docId}`, link])).values()];
  return <section className="document-backlinks" aria-label="双向链接"><div className="backlinks-heading"><h2>双向链接</h2><button aria-expanded={open} onClick={() => setOpen(value => !value)}>{open ? '收起' : `显示（${incoming.length}）`}</button></div>{open && <div className="backlinks-content">
    <h3>引用此文档</h3>{workspace.indexing && <p className="panel-empty" role="status">正在建立引用索引…</p>}{workspace.indexErrors.size > 0 && <p className="panel-empty">部分文档索引失败，引用列表可能不完整。</p>}{!incoming.length && !workspace.indexing && <p className="panel-empty">暂无反向链接。</p>}{incoming.map(link => <button data-peek-doc={link.sourceId} data-peek-workspace={workspace.id} className="backlink-row" key={link.sourceId} onClick={() => onNavigate(link.sourceId, workspace.id)}><span>{workspace.documents.get(link.sourceId)?.title || '无标题文档'} <small>{link.count > 1 ? `${link.count} 处引用` : ''}</small></span><small>{link.preview}</small></button>)}
    <h3>本文引用</h3>{!outgoing.length && <p className="panel-empty">尚未引用其他文档。输入 [[ 可选择文档。</p>}{outgoing.map(link => { const target = link.workspaceId === workspace.id ? workspace.documents.get(link.docId) : undefined; const missing = link.workspaceId === workspace.id && !target; return <button data-peek-doc={link.docId} data-peek-workspace={link.workspaceId} className="outgoing-reference" key={`${link.workspaceId}:${link.docId}`} disabled={missing} onClick={() => onNavigate(link.docId, link.workspaceId)}>{target?.title.trim() || link.label || '无标题文档'}{missing ? '（已删除）' : ''}</button>; })}
  </div>}</section>;
}
