import {recordNavigation,restoreNavigation,type NavigationHistory} from './navigation';
export type PaneId = 'left' | 'right';
export interface WorkbenchTab { docId: string; mode: 'page' | 'board'; frameId?: string; tick?: number }
export interface WorkbenchPaneState { tabs: WorkbenchTab[]; active: string | null; history?:NavigationHistory }
export interface WorkbenchState { version: 1; split: boolean; activePane: PaneId; ratio: number; left: WorkbenchPaneState; right: WorkbenchPaneState }
export const emptyWorkbench = (): WorkbenchState => ({ version: 1, split: false, activePane: 'left', ratio: 50, left: { tabs: [], active: null }, right: { tabs: [], active: null } });
export const activeTab = (state: WorkbenchState, pane = state.activePane) => state[pane].tabs.find(tab => tab.docId === state[pane].active);
export function openTab(state: WorkbenchState, docId: string, pane = state.activePane, frameId?: string, record=true): WorkbenchState {
  const target = state[pane]; const found = target.tabs.find(tab => tab.docId === docId); const next = frameId ? { docId, mode: 'board' as const, frameId, tick: Date.now() } : found ?? { docId, mode: 'page' as const };
  const tabs = found ? target.tabs.map(tab => tab.docId === docId ? next : tab) : [...target.tabs, next];
  return { ...state, split: state.split || pane === 'right', activePane: pane, [pane]: { ...target,tabs, active: docId,...(record?{history:recordNavigation(target.history,next)}:{}) } };
}
export function closeTab(state: WorkbenchState, pane: PaneId, docId: string): WorkbenchState {
  const target = state[pane], index = target.tabs.findIndex(tab => tab.docId === docId); const tabs = target.tabs.filter(tab => tab.docId !== docId);
  const active=target.active === docId ? tabs[Math.min(index, tabs.length - 1)]?.docId ?? null : target.active;const next=tabs.find(tab=>tab.docId===active);return { ...state, [pane]: { ...target,tabs,active,...(next&&active!==target.active?{history:recordNavigation(target.history,next)}:{}) } };
}
export function reconcileWorkbench(state: WorkbenchState, ids: Set<string>): WorkbenchState {
  let next = state; for (const pane of ['left','right'] as const) for (const tab of state[pane].tabs) if (!ids.has(tab.docId)) next = closeTab(next, pane, tab.docId); for(const pane of ['left','right'] as const){const history=next[pane].history;if(history?.entries.some(entry=>!ids.has(entry.docId)))next={...next,[pane]:{...next[pane],history:restoreNavigation(history,ids)}};}return next;
}
export function restoreWorkbench(value: unknown, ids: Set<string>): WorkbenchState {
  const fallback = emptyWorkbench(); if (!value || typeof value !== 'object') return fallback; const v = value as Partial<WorkbenchState>;
  if (v.version !== 1 || typeof v.split !== 'boolean' || !['left','right'].includes(String(v.activePane)) || typeof v.ratio !== 'number' || !Number.isFinite(v.ratio)) return fallback;
  const parse = (input: unknown): WorkbenchPaneState | null => { if (!input || typeof input !== 'object') return null; const p = input as WorkbenchPaneState; if (!Array.isArray(p.tabs) || p.tabs.length > 10000 || (p.active !== null && typeof p.active !== 'string')) return null; const seen = new Set<string>(); const tabs: WorkbenchTab[] = []; for (const tab of p.tabs) { if (!tab || typeof tab.docId !== 'string' || !['page','board'].includes(tab.mode) || (tab.frameId !== undefined && typeof tab.frameId !== 'string')) return null; if (ids.has(tab.docId) && !seen.has(tab.docId)) { seen.add(tab.docId); tabs.push({ docId: tab.docId, mode: tab.mode, ...(tab.frameId && tab.mode==='board' ? {frameId:tab.frameId}: {}) }); } } const active=tabs.some(t => t.docId === p.active) ? p.active : tabs[0]?.docId ?? null;let history=restoreNavigation(p.history,ids);const selected=tabs.find(tab=>tab.docId===active);if(selected)history=recordNavigation(history,selected);return {tabs,active,history}; };
  const left = parse(v.left), right = parse(v.right); return left && right ? { version: 1, split: v.split, activePane: v.split ? v.activePane! : 'left', ratio: Math.max(30,Math.min(70,v.ratio)), left, right } : fallback;
}
export function mergePanes(state: WorkbenchState): WorkbenchState {
  const active = activeTab(state); const tabs = [...state.left.tabs, ...state.right.tabs.filter(tab => !state.left.tabs.some(left => left.docId === tab.docId))].map(tab=>active?.docId===tab.docId?active:tab);
  return { ...state, split: false, activePane: 'left', left: { tabs, active: active?.docId ?? tabs[0]?.docId ?? null,history:active?recordNavigation(state[state.activePane].history,active):state.left.history }, right: { tabs: [], active: null } };
}

export function changeTabMode(state:WorkbenchState,pane:PaneId,docId:string,mode:'page'|'board'):WorkbenchState {const tab=state[pane].tabs.find(tab=>tab.docId===docId);if(!tab)return state;const next={docId,mode};return {...state,[pane]:{...state[pane],tabs:state[pane].tabs.map(tab=>tab.docId===docId?next:tab),history:recordNavigation(state[pane].history,next)}};}
export function navigateHistory(state:WorkbenchState,pane:PaneId,cursor:number):WorkbenchState {const history=state[pane].history,entry=history?.entries[cursor];if(!history||!entry)return state;let next=openTab(state,entry.docId,pane,undefined,false);const tab={...entry,...(entry.frameId?{tick:Date.now()}:{})};return {...next,[pane]:{...next[pane],tabs:next[pane].tabs.map(current=>current.docId===entry.docId?tab:current),history:{...history,cursor}}};}
