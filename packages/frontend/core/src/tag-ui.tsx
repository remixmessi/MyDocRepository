import { Check } from 'lucide-react';
import { TAG_COLORS, type TagColor, type TagMeta } from './organization';
import './tag-styles.css';

export function TagBadge({ tag }: { tag: TagMeta }) {
  return <span className="tag-badge" data-tag-color={tag.color ?? 'default'}>{tag.name}</span>;
}

export function TagColorPicker({ label, value, onChange, disabled }: {
  label: string; value: TagColor; onChange: (color: TagColor) => void; disabled?: boolean;
}) {
  return <div className="tag-color-picker" role="group" aria-label={label}>
    <span className="tag-color-label" aria-hidden="true">颜色</span>
    {TAG_COLORS.map(color => <button key={color.id} type="button" className="tag-color-swatch"
      data-tag-color={color.id} aria-label={color.label} title={color.label}
      aria-pressed={value === color.id} disabled={disabled} onClick={() => onChange(color.id)}>
      {value === color.id ? <Check size={16} aria-hidden="true" /> : <span className="tag-color-dot" aria-hidden="true" />}
    </button>)}
  </div>;
}
