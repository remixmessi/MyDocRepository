import {footnoteIndex} from './footnote-format';
import{calloutLabels}from'./expression-format';
import{paintExpression}from'./expression-render';
import type { Node as PMNode } from 'prosemirror-model';
import { imageMime, type WorkspaceAssets } from '@mydoc/nbstore';
import { webUrl } from './web-links';
export function renderEmbedPreview(doc: PMNode, host: HTMLElement, assets: WorkspaceAssets | undefined, title: (id: string, workspace: string) => string | undefined, navigate: (id: string, workspace: string) => void) {
  const footnotes=footnoteIndex(doc);
  const expressions:(()=>void)[]=[];let active = true; const urls = new Set<string>(); let remaining = 250; let textLeft = 10000; let truncated = false; host.replaceChildren();
  const text = (value: string) => { if (value.length > textLeft) truncated = true; const result = value.slice(0, textLeft); textLeft -= result.length; return document.createTextNode(result); };
  const reference = (node: PMNode, prefix = '') => { const link = document.createElement('button'); link.type = 'button'; link.className = 'preview-reference'; link.textContent = prefix + (title(node.attrs.docId, node.attrs.workspaceId) || node.attrs.label || '引用文档'); link.addEventListener('click', () => navigate(node.attrs.docId, node.attrs.workspaceId)); return link; };
  function render(node: PMNode): Node | null {
    if (remaining-- <= 0 || textLeft <= 0) { truncated = true; return null; }
    if (node.isText) { let result: Node = text(node.text || ''); for (const mark of node.marks) {
      const tags: Record<string,string> = { strong: 'strong', em: 'em', code: 'code', strike: 's' }; let wrapper: HTMLElement | null = null;
      if (tags[mark.type.name]) wrapper = document.createElement(tags[mark.type.name]);
      if (mark.type.name === 'link') { try { const link = document.createElement('a'); link.href = webUrl(mark.attrs.href); link.target = '_blank'; link.rel = 'noopener noreferrer'; wrapper = link; } catch { /* unsafe and internal plain links remain text */ } }
      if (wrapper) { wrapper.append(result); result = wrapper; }
    } return result; }
    if(['math_inline','math_block','mermaid'].includes(node.type.name)){const value=document.createElement(node.type.name==='math_inline'?'span':'div');expressions.push(paintExpression(value,node.type.name,node.attrs.source));return value;}
    if(node.type.name==='footnote_reference'){const sup=document.createElement('sup');sup.textContent='['+(footnotes.numbers.get(node.attrs.label)??'?')+']';return sup;}
    if (node.type.name === 'doc_reference') return reference(node, '↗ ');
    if (node.type.name === 'doc_embed') return reference(node, '嵌入文档：');
    if (node.type.name === 'image') {
      const wrapper = document.createElement('span'); wrapper.className = 'preview-image'; const name = node.attrs.alt || '图片'; const label = document.createElement('small'); label.textContent = name; wrapper.append(label);
      const id = String(node.attrs.src).startsWith('asset:') ? node.attrs.src.slice(6) : null;
      if (id && assets) void assets.get(id).then(record => { if (!active) return; const mime = record && imageMime(record.bytes); if (!record || !mime) { label.textContent = `图片不可用：${name}`; return; } const img = document.createElement('img'); const url = URL.createObjectURL(new Blob([record.bytes], { type: mime })); urls.add(url); img.src = url; img.alt = name; img.addEventListener('error', () => { if (active) label.textContent = `图片无法显示：${name}`; }); wrapper.prepend(img); }).catch(() => { if (active) label.textContent = `图片读取失败：${name}`; });
      else if (!id) { try { const img = document.createElement('img'); img.src = webUrl(node.attrs.src); img.alt = name; wrapper.prepend(img); } catch { label.textContent = `图片不可用：${name}`; } } return wrapper;
    }
    if (node.type.name === 'attachment') { const value = document.createElement('span'); value.className = 'preview-attachment'; value.append(text(`附件：${node.attrs.name}（打开原文档下载）`)); return value; }
    if (node.type.name === 'web_card') { const value = document.createElement('div'); value.className = 'preview-web-card'; value.append(text(`${node.attrs.title || node.attrs.url} ${node.attrs.description}`)); return value; }
    const tags: Record<string,string> = { paragraph:'p', blockquote:'blockquote', bullet_list:'ul', ordered_list:'ol', list_item:'li', code_block:'pre', horizontal_rule:'hr', hard_break:'br', table:'table', table_row:'tr', table_cell:'td', table_header:'th' };
    const element = document.createElement(node.type.name==='callout'?'aside':node.type.name === 'heading' ? `h${Math.min(6, Math.max(1, Number(node.attrs.level) || 1))}` : tags[node.type.name] || 'div');
    if(node.type.name==='footnote_definition'){element.className='footnote-definition';const heading=document.createElement('strong');heading.textContent=footnotes.numbers.has(node.attrs.label)?'脚注 '+footnotes.numbers.get(node.attrs.label):'脚注（未引用）';element.append(heading);}
    if(node.type.name==='callout'){element.className='doc-callout';element.dataset.callout=node.attrs.kind;element.dataset.calloutLabel=calloutLabels[node.attrs.kind as 'note'];element.setAttribute('role','note');element.setAttribute('aria-label',element.dataset.calloutLabel);}
    if (node.type.name === 'ordered_list') element.setAttribute('start', String(Number(node.attrs.order) || 1));
    let content = element;
    if (node.type.name === 'list_item' && node.attrs.checked !== null) { element.classList.add('preview-task'); const check = document.createElement('span'); check.className = 'preview-task-check'; check.textContent = node.attrs.checked ? '☑' : '☐'; content = document.createElement('div'); content.className = 'preview-task-content'; element.append(check, content); }
    if (['table_cell','table_header'].includes(node.type.name) && ['left','center','right'].includes(node.attrs.align)) element.style.textAlign = node.attrs.align;
    node.forEach(child => { const result = render(child); if (result) content.append(result); }); return element;
  }
  if (doc.childCount === 1 && doc.firstChild?.type.name === 'paragraph' && !doc.firstChild.content.size) { const empty = document.createElement('p'); empty.className = 'embed-empty'; empty.textContent = '原文档还没有正文。'; host.append(empty); return () => { active = false; }; }
  doc.forEach(node => { const result = render(node); if (result) host.append(result); });
  if (!host.childNodes.length) { const empty = document.createElement('p'); empty.className = 'embed-empty'; empty.textContent = '原文档还没有正文。'; host.append(empty); }
  if (truncated) { const hint = document.createElement('p'); hint.className = 'embed-empty'; hint.textContent = '预览已截断，请打开原文档查看完整内容。'; host.append(hint); }
  return () => { active = false;expressions.forEach(clean=>clean()); urls.forEach(url => URL.revokeObjectURL(url)); urls.clear(); };
}
