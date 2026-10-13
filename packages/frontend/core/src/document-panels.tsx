import { useEffect, useId, useRef, useState, type CSSProperties } from 'react';
import type { OutlineEntry } from '@mydoc/editor-blocks';
import type { LocalWorkspace } from './workspace';

export function DocumentOutlineRail({ entries, onJump }: { entries: OutlineEntry[]; onJump: (position: number) => void }) {
  const rail = useRef<HTMLElement>(null);
  const list = useRef<HTMLElement>(null);
  const listId = useId();
  const [open, setOpen] = useState(false);
  const [hovered, setHovered] = useState<number>();
  const [bounds, setBounds] = useState({ right: 12, top: 120, height: 300, markerHeight: 62, popoverTop: 0, width: 260, visible: false });
  const signature = JSON.stringify(entries);
  useEffect(() => {
    const host = rail.current;
    if (!host || !entries.length) return;
    const layout = host.closest('.document-layout');
    const root = layout?.querySelector<HTMLElement>('.document-editor .ProseMirror');
    const pane = host.closest<HTMLElement>('.workbench-content');
    if (!root || !pane) return;
    let frame = 0;
    const measure = () => {
      frame = 0;
      const port = pane.getBoundingClientRect();
      const header = pane.querySelector('.editor-header')?.getBoundingClientRect();
      // Anchor to the document frame, independent of the centered text column.
      const contentTop = Math.max(24, port.top + 12, (header?.bottom ?? 0) + 20);
      const bottom = Math.min(innerHeight - 24, port.bottom - 12);
      const height = Math.max(0, bottom - contentTop);
      const markerHeight = Math.min(height, entries.length * 18 + 8);
      const top = contentTop + (height - markerHeight) / 2;
      const rightEdge = port.right - 12;
      const popoverHeight = list.current?.offsetHeight || Math.min(height, entries.length * 40 + 12);
      const popoverTop = Math.max(contentTop - top, Math.min(0, bottom - top - popoverHeight));
      const nextBounds = { right: Math.max(12, innerWidth - rightEdge), top, height, markerHeight, popoverTop,
        width: Math.min(300, Math.max(120, rightEdge - 42 - Math.max(0, port.left) - 12)),
        visible: root.getClientRects().length > 0 && height > 40 && port.width > 160 };
      setBounds(previous => Object.keys(nextBounds).every(key => previous[key as keyof typeof previous] === nextBounds[key as keyof typeof nextBounds]) ? previous : nextBounds);
    };
    const schedule = () => { if (!frame) frame = requestAnimationFrame(measure); };
    const resize = new ResizeObserver(schedule);
    resize.observe(root); resize.observe(pane);
    if (list.current) resize.observe(list.current);
    const mutations = new MutationObserver(schedule);
    mutations.observe(root, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['hidden', 'class'] });
    if (layout) mutations.observe(layout, { subtree: true, attributes: true, attributeFilter: ['hidden'] });
    document.addEventListener('scroll', schedule, { passive: true, capture: true });
    window.addEventListener('resize', schedule);
    schedule();
    return () => {
      if (frame) cancelAnimationFrame(frame);
      resize.disconnect(); mutations.disconnect();
      document.removeEventListener('scroll', schedule, true);
      window.removeEventListener('resize', schedule);
    };
  }, [signature]);
  const reveal = (index: number) => {
    setHovered(index); setOpen(true);
    requestAnimationFrame(() => list.current?.querySelectorAll('button')[index]?.scrollIntoView({ block: 'nearest' }));
  };
  const jump = (position: number) => { setOpen(false); onJump(position); };
  if (!entries.length) return null;
  return <aside ref={rail} className="document-outline-rail" aria-label="标题刻度" hidden={!bounds.visible}
    style={{ right: bounds.right, top: bounds.top, maxHeight: bounds.markerHeight, '--outline-popover-width': `${bounds.width}px` } as CSSProperties}
    onMouseLeave={() => { setHovered(undefined); if (!rail.current?.contains(document.activeElement)) setOpen(false); }}
    onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) { setOpen(false); setHovered(undefined); } }}
    onKeyDown={event => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); setOpen(false); } }}>
    <div className="document-outline-markers" style={{ maxHeight: bounds.markerHeight }}>
      {entries.map((entry, index) => <button type="button" key={entry.position}
        className={`document-outline-marker${open && index === hovered ? ' highlighted' : ''}`}
        aria-label={`跳转到标题：${entry.text}`}
        aria-expanded={open} aria-controls={listId}
        onMouseEnter={() => reveal(index)} onFocus={() => reveal(index)} onClick={() => jump(entry.position)}>
        <span className="document-outline-marker-line" aria-hidden="true" style={{ width: 32 - (entry.level - 1) * 4 }} />
      </button>)}
    </div>
    <nav ref={list} id={listId} className="document-outline-popover" aria-label="悬浮大纲" hidden={!open} style={{ top: bounds.popoverTop, maxHeight: bounds.height }}>
      {entries.map((entry, index) => <button type="button" key={entry.position}
        className={index === hovered ? 'hovered' : undefined}
        style={{ paddingLeft: 12 + (entry.level - 1) * 12 }} onMouseEnter={() => setHovered(index)} onFocus={() => setHovered(index)} onClick={() => jump(entry.position)}>{entry.text}</button>)}
    </nav>
  </aside>;
}

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
