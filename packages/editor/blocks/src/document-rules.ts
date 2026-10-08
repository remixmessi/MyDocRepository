import{webCardPositionAllowed}from'./web-links';
import { InputRule, inputRules, wrappingInputRule, textblockTypeInputRule } from 'prosemirror-inputrules';
import { Plugin, TextSelection } from 'prosemirror-state';
import { documentSchema as schema } from './document-format';

export function documentInputRules() {
  const markRules = [[/\*\*([^*]+)\*\*$/, 'strong'], [/__([^_]+)__$/, 'strong'], [/(?<!\*)\*([^*]+)\*$/, 'em'], [/`([^`]+)`$/, 'code'], [/~~([^~]+)~~$/, 'strike']] as const;
  return inputRules({ rules: [
    new InputRule(/(?:^|\s)\$([^$\n]+)\$$/,(state,match,start,end)=>{const source=match[1];if(/^\s|\s$/.test(source))return null;const from=end-source.length-1;return state.tr.replaceWith(from,end,schema.nodes.math_inline.create({source}));},{inCode:false}),
    textblockTypeInputRule(/^(#{1,6})\s$/, schema.nodes.heading, match => ({ level: match[1].length })),
    wrappingInputRule(/^>\s$/, schema.nodes.blockquote),
    wrappingInputRule(/^[-+*]\s$/, schema.nodes.bullet_list, { tight: true }),
    wrappingInputRule(/^(\d+)\.\s$/, schema.nodes.ordered_list, match => ({ order: +match[1], tight: true }), (match, node) => node.childCount + node.attrs.order === +match[1]),
    new InputRule(/^\[([ xX])\]\s$/, (state, match, start, end) => {
      const $from = state.selection.$from;
      for (let depth = $from.depth; depth > 0; depth--) {
        if ($from.node(depth).type === schema.nodes.list_item) return state.tr.delete(start, end).setNodeMarkup($from.before(depth), undefined, { checked: match[1].toLowerCase() === 'x' });
      }
      return null;
    }),
    ...markRules.map(([pattern, name]) => new InputRule(pattern, (state, match, start, end) => {
      const content = match[1]; const tr = state.tr.insertText(content, start, end).addMark(start, start + content.length, schema.marks[name].create());
      tr.setSelection(TextSelection.create(tr.doc, start + content.length)); tr.removeStoredMark(schema.marks[name]); return tr;
    }, { inCode: false })),
    new InputRule(/\[([^\]]+)\]\((https?:\/\/[^\s)]+|mailto:[^\s)]+)\)$/, (state, match, start, end) => {
      const tr = state.tr.insertText(match[1], start, end).addMark(start, start + match[1].length, schema.marks.link.create({ href: match[2] }));
      tr.setSelection(TextSelection.create(tr.doc, start + match[1].length)); tr.removeStoredMark(schema.marks.link); return tr;
    }),
  ] });
}
export function markdownEnter(state: import('prosemirror-state').EditorState, dispatch?: (tr: import('prosemirror-state').Transaction) => void) {
  const { $from, empty } = state.selection; if (!empty) return false;
  const text = $from.parent.textContent;
  const blockAllowed=webCardPositionAllowed(state.selection);
  const alert=/^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]$/i.exec(text);if(alert&&blockAllowed&&$from.parent.type===schema.nodes.paragraph){for(let depth=$from.depth-1;depth>0;depth--)if($from.node(depth).type===schema.nodes.blockquote&&$from.node(depth).firstChild===$from.parent){dispatch?.(state.tr.delete($from.start(),$from.end()).setNodeMarkup($from.before(depth),schema.nodes.callout,{kind:alert[1].toLowerCase()}));return true;}}
  if ($from.parent.type === schema.nodes.paragraph) {
    if(text==='$$'&&blockAllowed){const pos=$from.before(),tr=state.tr.replaceWith(pos,$from.after(),[schema.nodes.math_block.create({source:''}),schema.nodes.paragraph.create()]);tr.setSelection(TextSelection.create(tr.doc,pos+2));dispatch?.(tr);return true;}
    const fence = /^```([\w-]*)$/.exec(text);
    if(fence?.[1]==='mermaid'&&blockAllowed){const pos=$from.before(),tr=state.tr.replaceWith(pos,$from.after(),[schema.nodes.mermaid.create({source:''}),schema.nodes.paragraph.create()]);tr.setSelection(TextSelection.create(tr.doc,pos+2));dispatch?.(tr);return true;}
    if (fence) { dispatch?.(state.tr.delete($from.start(), $from.end()).setNodeMarkup($from.before(), schema.nodes.code_block, { params: fence[1] })); return true; }
    if (/^(---|\*\*\*|___)$/.test(text)) {
      const position = $from.before(); const tr = state.tr.replaceWith(position, $from.after(), [schema.nodes.horizontal_rule.create(), schema.nodes.paragraph.create()]);
      tr.setSelection(TextSelection.create(tr.doc, position + 2)); dispatch?.(tr); return true;
    }
  }
  if ($from.parent.type === schema.nodes.code_block && $from.parentOffset === text.length && /(?:^|\n)```$/.test(text)) {
    const start = $from.start() + text.length - 3 - (text.length > 3 ? 1 : 0);
    const tr = state.tr.delete(start, $from.end()); const after = tr.mapping.map($from.after());
    tr.insert(after, schema.nodes.paragraph.create()).setSelection(TextSelection.create(tr.doc, after + 1)); dispatch?.(tr); return true;
  }
  return false;
}
export const compositionGuard = new Plugin({ props: { handleDOMEvents: { keydown: (_, event) => (event as KeyboardEvent).isComposing } } });
