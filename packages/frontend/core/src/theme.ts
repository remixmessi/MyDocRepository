import {projectPreferences} from '@mydoc/nbstore';
export type ThemeMode = 'light' | 'dark' | 'system';
export const THEME_KEY = 'mydoc:theme:v1';
export function themeMode(value: unknown): ThemeMode { return value==='light'||value==='dark'||value==='system'?value:'system'; }
export function resolvedTheme(mode: ThemeMode, darkSystem: boolean) { return mode==='system' ? darkSystem?'dark':'light' : mode; }
export function readTheme(storage: Pick<Storage,'getItem'>): ThemeMode { try { return themeMode(storage.getItem(THEME_KEY)); } catch { return 'system'; } }
export function saveTheme(storage: Pick<Storage,'setItem'>, mode: ThemeMode) { storage.setItem(THEME_KEY,mode); }
export function applyTheme(mode: ThemeMode, darkSystem=matchMedia('(prefers-color-scheme: dark)').matches) {
  const value=resolvedTheme(mode,darkSystem),root=document.documentElement;
  if(root.dataset.theme!==value){root.dataset.theme=value;window.dispatchEvent(new CustomEvent('mydoc-theme-change'));}
}
export function initializeTheme(){let mode:ThemeMode='system';try{mode=readTheme(projectPreferences);}catch{/* Use the system appearance when browser storage is unavailable. */}applyTheme(mode);return mode;}
