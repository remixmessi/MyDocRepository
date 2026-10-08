import Prism from 'prismjs';
import 'prismjs/components/prism-typescript';
import 'prismjs/components/prism-sql';
import 'prismjs/components/prism-json';
import 'prismjs/components/prism-python';
import 'prismjs/components/prism-bash';
import 'prismjs/components/prism-java';
import 'prismjs/components/prism-c';
import 'prismjs/components/prism-cpp';
import 'prismjs/components/prism-csharp';
import 'prismjs/components/prism-go';
import 'prismjs/components/prism-rust';
import 'prismjs/components/prism-yaml';
import 'prismjs/components/prism-markdown';
import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node } from 'prosemirror-model';
function decorate(doc: Node) {
  const decorations: Decoration[] = [];
  doc.descendants((node, position) => {
    if (node.type.name !== 'code_block') return;
    const aliases: Record<string, string> = { js: 'javascript', ts: 'typescript', html: 'markup' };
    const language = aliases[node.attrs.params] ?? node.attrs.params;
    const grammar = Prism.languages[language]; if (!grammar) return false;
    let offset = position + 1;
    const visit = (tokens: (string | Prism.Token)[]) => tokens.forEach(token => {
      if (typeof token === 'string') { offset += token.length; return; }
      const start = offset;
      if (typeof token.content === 'string') offset += token.content.length;
      else visit(Array.isArray(token.content) ? token.content : [token.content]);
      if (offset > start) decorations.push(Decoration.inline(start, offset, { class: `token ${token.type}` }));
    });
    visit(Prism.tokenize(node.textContent, grammar)); return false;
  });
  return DecorationSet.create(doc, decorations);
}
export const codeHighlight = new Plugin({ state: { init: (_, state) => decorate(state.doc), apply: (tr, previous) => tr.docChanged ? decorate(tr.doc) : previous }, props: { decorations(state) { return this.getState(state); } } });
