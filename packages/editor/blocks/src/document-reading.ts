import {projectPreferences} from '@mydoc/nbstore';
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { ySyncPluginKey, relativePositionToAbsolutePosition } from 'y-prosemirror';
import { Y } from '@mydoc/crdt';
import type { BlockStore } from '@mydoc/editor-framework';
import { createCommentAnchor, type CommentMapping } from './comment-model';
import { ReadingPositions, readingKey, boundedScroll, readingSaveError, type ReadingPosition } from './reading-position';

export function readingJump(view: EditorView) {
  view.dom.dispatchEvent(new CustomEvent('mydoc-reading-jump', { bubbles: true }));
}

export function documentReading(store: BlockStore, workspaceId: string, documentId: string, notify: (error: string) => void) {
  let retry = () => {};
  return { retry: () => retry(), plugin: new Plugin({ view(view) {
    let disposed = false, frame = 0, timer = 0, release = 0, adjustment = 0;
    let adjusting = false, suspended = false, restoring = true, interacting = false, failed = false;
    let last: ReadingPosition | undefined, positions: ReadingPositions | undefined;
    const host = view.dom.closest<HTMLElement>('.workbench-content');
    const pane = host?.id === 'pane-right' ? 'right' : 'left';
    const key = readingKey(workspaceId, documentId, pane);
    try { positions = new ReadingPositions(projectPreferences); last = positions.get(key); }
    catch { /* Unavailable browser storage must not prevent editing. */ }
    const root = view.dom.closest<HTMLElement>('.markdown-document') ?? view.dom;
    const fragment = store.doc.getXmlFragment('document');
    const rootOverflow = root.style.overflowAnchor, bodyOverflow = view.dom.style.overflowAnchor;
    // Split panes scroll independently; narrow screens and ordinary pages use the window.
    const scroller = () => host && ['auto', 'scroll'].includes(getComputedStyle(host).overflowY)
      ? host : document.scrollingElement as HTMLElement;
    const port = () => scroller() === document.scrollingElement
      ? { top: 0, bottom: innerHeight } : scroller().getBoundingClientRect();
    const mapping = () => ySyncPluginKey.getState(view.state)?.binding?.mapping as CommentMapping | undefined;
    const visible = () => view.dom.getClientRects().length > 0 && root.getClientRects().length > 0 && !root.hidden;
    const save = async () => {
      clearTimeout(timer);
      if (!last || disposed) return;
      try {
        if (!positions) positions = new ReadingPositions(projectPreferences);
        await positions.put(key, last);
        if (failed) { failed = false; notify(''); }
      } catch (error) {
        failed = true;
        notify(readingSaveError(error));
      }
    };
    retry = save;
    const capture = () => {
      if (disposed || !visible() || document.querySelector('dialog[open]')) return;
      const scroll = scroller(), max = Math.max(0, scroll.scrollHeight - scroll.clientHeight);
      const scrollTop = boundedScroll(scroll.scrollTop, max);
      const bounds = view.dom.getBoundingClientRect(), viewport = port();
      let anchor: ReadingPosition['anchor'] = null, offset = 0;
      const top = Math.max(viewport.top + 24, bounds.top + 2);
      if (top < Math.min(viewport.bottom, bounds.bottom)) {
        const found = view.posAtCoords({ left: Math.max(bounds.left + 8, Math.min(bounds.left + 32, bounds.right - 8)), top });
        const map = mapping();
        if (found && map) {
          try {
            anchor = createCommentAnchor(store.doc, fragment, map, found.pos, 0);
            offset = view.coordsAtPos(found.pos).top - viewport.top;
          } catch { /* Preserve the pixel fallback while the document mapping updates. */ }
        }
      }
      last = { version: 1, anchor, offset, scroll: scrollTop, ratio: max > 0 ? scrollTop / max : 0, updatedAt: Date.now() };
      clearTimeout(timer);
      timer = window.setTimeout(save, 400);
    };
    const compensate = () => {
      if (disposed || !visible() || !last || suspended || interacting || view.composing || document.querySelector('dialog[open]')) return;
      const scroll = scroller(), map = mapping(), max = scroll.scrollHeight - scroll.clientHeight;
      let target = last.scroll;
      // A reader at the beginning should keep the title visible when its height changes.
      if (last.scroll <= 1 && !last.ratio) target = 0;
      else if (last.anchor && map) {
        try {
          const pos = relativePositionToAbsolutePosition(store.doc, fragment, Y.createRelativePositionFromJSON(last.anchor), map);
          target = pos !== null && pos >= 0 && pos <= view.state.doc.content.size
            ? scroll.scrollTop + view.coordsAtPos(pos).top - port().top - last.offset : last.ratio * max;
        } catch { target = last.ratio * max; }
      }
      const next = boundedScroll(target, max);
      if (Math.abs(next - scroll.scrollTop) > .75) {
        adjusting = true;
        scroll.scrollTop = next;
        cancelAnimationFrame(adjustment);
        adjustment = requestAnimationFrame(() => { adjusting = false; });
      }
    };
    const schedule = () => {
      if (!frame) frame = requestAnimationFrame(() => {
        frame = 0;
        if (restoring) {
          if (last) compensate(); else scroller().scrollTop = 0;
          restoring = false;
          if (!last) capture();
        } else compensate();
      });
    };
    const settle = () => {
      clearTimeout(release);
      release = window.setTimeout(() => { suspended = false; interacting = false; capture(); }, 160);
    };
    const scroll = (event: Event) => {
      const current = scroller();
      if (event.target !== current && !(current === document.scrollingElement && event.target === document)) return;
      // Layout-induced scroll events must not replace the anchor before compensation.
      if (disposed || adjusting || restoring || (frame && !suspended)) return;
      capture();
    };
    const intent = (event: Event) => {
      const target = event.target as HTMLElement;
      if (!view.dom.contains(target) && event.type !== 'wheel' && event.type !== 'touchstart') return;
      if (target?.closest('dialog,.document-comments-panel,.reference-popup,.slash-menu,.document-outline')) return;
      suspended = true; interacting = true; restoring = false;
      settle();
    };
    const jump = () => {
      suspended = true; interacting = false; restoring = false;
      clearTimeout(release);
      release = window.setTimeout(() => { suspended = false; capture(); }, 180);
    };
    const resize = () => schedule();
    const observed = new Set<Element>();
    const size = new ResizeObserver(resize);
    const observeBlocks = () => {
      const elements = new Set<Element>([root, ...view.dom.children]);
      for (const element of observed) if (!elements.has(element)) { size.unobserve(element); observed.delete(element); }
      for (const element of elements) if (!observed.has(element)) { observed.add(element); size.observe(element); }
    };
    const mutation = new MutationObserver(() => { observeBlocks(); schedule(); });
    // One controller owns compensation instead of competing with native scroll anchoring.
    root.style.overflowAnchor = 'none'; view.dom.style.overflowAnchor = 'none';
    observeBlocks(); mutation.observe(root, { subtree: true, childList: true });
    document.addEventListener('scroll', scroll, true);
    root.addEventListener('mydoc-reading-jump', jump);
    root.addEventListener('wheel', intent, { passive: true });
    root.addEventListener('touchstart', intent, { passive: true });
    root.addEventListener('pointerdown', intent); root.addEventListener('keydown', intent);
    window.addEventListener('resize', resize); window.addEventListener('pagehide', save);window.addEventListener('mydoc-before-close',save);
    document.addEventListener('visibilitychange', save);
    schedule();
    return {
      update(next, previous) { view = next; if (next.state.doc !== previous.doc) schedule(); },
      destroy() {
        // React StrictMode may clean up before the first restoration frame.
        if (!restoring) capture();
        save(); disposed = true; retry = () => {};
        cancelAnimationFrame(frame); cancelAnimationFrame(adjustment);
        clearTimeout(timer); clearTimeout(release); mutation.disconnect(); size.disconnect();
        document.removeEventListener('scroll', scroll, true);
        root.removeEventListener('mydoc-reading-jump', jump);
        root.removeEventListener('wheel', intent); root.removeEventListener('touchstart', intent);
        root.removeEventListener('pointerdown', intent); root.removeEventListener('keydown', intent);
        window.removeEventListener('resize', resize); window.removeEventListener('pagehide', save);window.removeEventListener('mydoc-before-close',save);
        document.removeEventListener('visibilitychange', save);
        root.style.overflowAnchor = rootOverflow; view.dom.style.overflowAnchor = bodyOverflow;
      },
    };
  } }) };
}
