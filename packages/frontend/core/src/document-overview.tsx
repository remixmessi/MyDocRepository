import {FileTypeIcon} from './file-type-icon';
import { useMemo, useState } from 'react';
import { FileText, Grid2X2, List, MoreHorizontal, Plus, Star } from 'lucide-react';
import type { LocalWorkspace, DocumentMeta } from './workspace';
import { snippet } from './search';
import { TagBadge } from './tag-ui';

type ViewMode = 'list' | 'grid';

interface Props {
  workspace: LocalWorkspace;
  busy: boolean;
  onCreate: () => void;
  onSelect: (id: string) => void;
  onFavorite: (id: string) => void;
  onAction: (id: string, kind: 'rename' | 'move' | 'delete') => void;
}

function dateLabel(timestamp: number) {
  return new Intl.DateTimeFormat('zh-CN', { year: 'numeric', month: 'long', day: 'numeric' }).format(timestamp);
}

function timeLabel(timestamp: number) {
  return new Intl.DateTimeFormat('zh-CN', { month: 'numeric', day: 'numeric' }).format(timestamp);
}

function titleFor(meta: DocumentMeta) {
  return meta.title.trim() || '无标题文档';
}

export function DocumentOverview({ workspace, busy, onCreate, onSelect, onFavorite, onAction }: Props) {
  const [view, setView] = useState<ViewMode>('list');
  const [section, setSection] = useState<'documents' | 'featured' | 'tags'>('documents');
  const [favoriteOnly, setFavoriteOnly] = useState(false);
  const [tagId, setTagId] = useState('');
  const tag = tagId && workspace.organization.tags.has(tagId) ? tagId : '';
  const documents = workspace.files()
    .filter(meta => (!favoriteOnly || !!meta.favorite) && (!tag || meta.tagIds?.includes(tag)))
    .sort((a, b) => b.createdAt - a.createdAt || b.id.localeCompare(a.id));
  const groups = useMemo(() => {
    const grouped = new Map<string, DocumentMeta[]>();
    for (const meta of documents) {
      const key = new Date(meta.createdAt).toLocaleDateString('zh-CN');
      const entries = grouped.get(key) ?? [];
      entries.push(meta);
      grouped.set(key, entries);
    }
    return [...grouped.entries()];
  }, [documents]);
  const tags = workspace.organization.tagList();

  return <section className="document-overview" aria-label="全部文档">
    <header className="overview-header">
      <div className="overview-heading">
        <div className="overview-tabs" role="tablist" aria-label="文档视图">
          <button role="tab" aria-selected={section === 'documents'} className={section === 'documents' ? 'active' : ''} onClick={() => { setSection('documents'); setFavoriteOnly(false); }}>文档</button>
          <button role="tab" aria-selected={section === 'featured'} className={section === 'featured' ? 'active' : ''} onClick={() => { setSection('featured'); setFavoriteOnly(true); }}>精选</button>
          <button role="tab" aria-selected={section === 'tags'} className={section === 'tags' ? 'active' : ''} onClick={() => { setSection('tags'); requestAnimationFrame(() => document.getElementById('overview-tag-filter')?.focus()); }}>标签</button>
        </div>
        <p>{documents.length} 篇文档</p>
      </div>
      <div className="overview-actions">
        <div className="overview-view-toggle" role="group" aria-label="视图布局">
          <button aria-label="列表视图" title="列表视图" aria-pressed={view === 'list'} onClick={() => setView('list')}><List size={16} /></button>
          <button aria-label="网格视图" title="网格视图" aria-pressed={view === 'grid'} onClick={() => setView('grid')}><Grid2X2 size={16} /></button>
        </div>
        <button className="overview-new" disabled={busy || workspace.readOnly} onClick={onCreate}><Plus size={16} />新建文档</button>
      </div>
    </header>
    <div className="overview-filters">
      <button className={favoriteOnly ? 'filter-active' : ''} aria-pressed={favoriteOnly} onClick={() => setFavoriteOnly(value => !value)}><Star size={14} />收藏</button>
      <label>标签<select id="overview-tag-filter" aria-label="按标签筛选" value={tagId} onChange={event => { setTagId(event.target.value); setSection('tags'); }} disabled={busy}><option value="">全部标签</option>{tags.map(tag => <option key={tag.id} value={tag.id}>{tag.name}</option>)}</select></label>
      {(favoriteOnly || tagId) && <button className="clear-overview-filter" onClick={() => { setFavoriteOnly(false); setTagId(''); }}>清除筛选</button>}
    </div>
    {!documents.length ? <div className="overview-empty"><FileText size={30} /><h2>{favoriteOnly || tagId ? '没有符合条件的文档' : '还没有文档'}</h2><p>{favoriteOnly || tagId ? '清除筛选后查看全部文档。' : '创建第一篇文档，开始记录内容。'}</p>{!favoriteOnly && !tagId && <button className="primary" disabled={busy || workspace.readOnly} onClick={onCreate}><Plus size={16} />新建文档</button>}</div> : <div className={`overview-groups ${view === 'grid' ? 'is-grid' : ''}`}>
      {groups.map(([key, entries]) => <section className="overview-group" key={key}>
        <h2>{dateLabel(entries[0].createdAt)} <span>{entries.length}</span></h2>
        <div className="overview-items">{entries.map(meta => {
          const title = titleFor(meta);
          const text = snippet(workspace.search.text(meta.id), '').trim();
          const itemTags = (meta.tagIds ?? []).map(id => workspace.organization.tags.get(id)).filter(Boolean);
          return <article className="overview-item" key={meta.id}>
            <button className="overview-open" disabled={busy} onClick={() => onSelect(meta.id)}>
              <span className="overview-icon"><FileTypeIcon resource={workspace.officeResources.get(meta.id)} size={18}/></span>
              <span className="overview-copy"><strong>{title}</strong><small>{text || '暂无正文预览'}</small></span>
              {itemTags.length > 0 && <span className="overview-tags">{itemTags.slice(0, 2).map(tag => <TagBadge key={tag!.id} tag={tag!} />)}</span>}
              <time dateTime={new Date(meta.createdAt).toISOString()}>{timeLabel(meta.createdAt)}</time>
            </button>
            <button className={`overview-star ${meta.favorite ? 'is-favorite' : ''}`} aria-label={meta.favorite ? `取消收藏 ${title}` : `收藏 ${title}`} aria-pressed={!!meta.favorite} disabled={busy || workspace.readOnly || workspace.documentReadOnly(meta.id)} onClick={() => onFavorite(meta.id)}><Star size={16} fill={meta.favorite ? 'currentColor' : 'none'} /></button>
            <details className="overview-menu"><summary aria-label={`文档操作 ${title}`}><MoreHorizontal size={17} /></summary><div role="menu"><button role="menuitem" onClick={() => onAction(meta.id, 'rename')}>重命名</button><button role="menuitem" onClick={() => onAction(meta.id, 'move')}>移动到文件夹</button><button role="menuitem" disabled={workspace.readOnly || workspace.documentReadOnly(meta.id)} onClick={() => onAction(meta.id, 'delete')}>删除</button></div></details>
          </article>;
        })}</div>
      </section>)}
    </div>}
  </section>;
}
