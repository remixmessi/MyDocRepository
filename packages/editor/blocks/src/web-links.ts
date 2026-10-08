import type { Selection } from 'prosemirror-state';
export type WebCardKind = 'bookmark' | 'embed';
export interface WebCardAttrs { id: string | null; kind: WebCardKind; url: string; title: string; description: string }
export function webUrl(value: string) {
  const url = new URL(value.trim()); if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) throw new Error('请输入不含账号密码的 HTTP 或 HTTPS 网址'); return url.href;
}
export function validWebCard(value: unknown): value is WebCardAttrs {
  if (!value || typeof value !== 'object') return false; const attrs = value as Record<string, unknown>;
  if ((attrs.id !== null && typeof attrs.id !== 'string') || !['bookmark', 'embed'].includes(String(attrs.kind)) || typeof attrs.url !== 'string' || typeof attrs.title !== 'string' || typeof attrs.description !== 'string' || attrs.url.length > 8192 || attrs.title.length > 500 || attrs.description.length > 2000) return false;
  try { webUrl(attrs.url); return true; } catch { return false; }
}
export function webCardMarkdown(attrs: WebCardAttrs) { return JSON.stringify(attrs); }
export function parseWebCardMarkdown(text: string): WebCardAttrs | null { try { const attrs: unknown = JSON.parse(text); return validWebCard(attrs) ? { ...attrs, url: webUrl(attrs.url) } : null; } catch { return null; } }
export async function readWebMetadata(value: string, signal: AbortSignal) {
  const response = await fetch(webUrl(value), { signal, credentials: 'omit', referrerPolicy: 'no-referrer' });
  if (!response.ok || !response.headers.get('content-type')?.includes('text/html')) throw new Error('网页没有提供可读取的 HTML 信息');
  if (!response.body) throw new Error('网页内容为空');
  const reader = response.body.getReader(); const chunks: Uint8Array[] = []; let size = 0;
  try { while (true) { const { done, value: bytes } = await reader.read(); if (done) break; size += bytes.byteLength; if (size > 1024 * 1024) throw new Error('网页内容过大，请手动填写标题和摘要'); chunks.push(bytes); } }
  finally { await reader.cancel().catch(() => {}); reader.releaseLock(); }
  const bytes = new Uint8Array(size); let offset = 0; for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
  const prefix = new TextDecoder().decode(bytes.subarray(0, 4096)); const charset = /charset\s*=\s*["']?([\w-]+)/i.exec(response.headers.get('content-type') || '')?.[1] || /<meta[^>]*charset\s*=\s*["']?([\w-]+)/i.exec(prefix)?.[1] || 'utf-8';
  let decoder: TextDecoder; try { decoder = new TextDecoder(charset); } catch { decoder = new TextDecoder(); } const html = decoder.decode(bytes);
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const clean = (text: string | null | undefined, max: number) => (text || '').replace(/\s+/g, ' ').trim().slice(0, max);
  return { title: clean(doc.querySelector('meta[property="og:title"]')?.getAttribute('content') || doc.querySelector('title')?.textContent, 500), description: clean(doc.querySelector('meta[property="og:description"]')?.getAttribute('content') || doc.querySelector('meta[name="description"]')?.getAttribute('content'), 2000) };
}

export function webCardPositionAllowed(selection: Selection) { const { $from } = selection; if ($from.parent.type.spec.code) return false; for (let depth = $from.depth; depth > 0; depth--) if (['table_cell', 'table_header'].includes($from.node(depth).type.name)) return false; return true; }
