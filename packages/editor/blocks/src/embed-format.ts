export interface DocEmbedAttrs { docId: string; workspaceId: string; label: string; collapsed: boolean }
export function validDocEmbed(value: unknown): value is DocEmbedAttrs {
  if (!value || typeof value !== 'object') return false; const attrs = value as Record<string, unknown>;
  return typeof attrs.docId === 'string' && !!attrs.docId && attrs.docId.length <= 200 && typeof attrs.workspaceId === 'string' && !!attrs.workspaceId && attrs.workspaceId.length <= 200 && typeof attrs.label === 'string' && attrs.label.length <= 500 && typeof attrs.collapsed === 'boolean';
}
export function parseDocEmbed(text: string): DocEmbedAttrs | null { try { const attrs: unknown = JSON.parse(text); return validDocEmbed(attrs) ? attrs : null; } catch { return null; } }
