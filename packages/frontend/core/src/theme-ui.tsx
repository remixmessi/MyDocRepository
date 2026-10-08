import {projectPreferences} from '@mydoc/nbstore';
import {useLayoutEffect,useState} from 'react';
import {Modal} from './dialogs';
import {applyTheme,initializeTheme,saveTheme,THEME_KEY,themeMode,type ThemeMode} from './theme';
export function useAppearance(){
 const [mode,setMode]=useState(initializeTheme);
 useLayoutEffect(()=>{applyTheme(mode);const system=matchMedia('(prefers-color-scheme: dark)'),change=()=>applyTheme(mode,system.matches),storage=(event:Event)=>{const detail=(event as CustomEvent<{key:string;value:string}>).detail;if(detail?.key===THEME_KEY)setMode(themeMode(detail.value));};system.addEventListener('change',change);window.addEventListener('mydoc-preference',storage);return()=>{system.removeEventListener('change',change);window.removeEventListener('mydoc-preference',storage);};},[mode]);
 return {mode,setMode,async save(value:ThemeMode){await projectPreferences.setItem(THEME_KEY,value);setMode(value);}};
}
export function ThemeDialog({initial,preview,save,onClose}:{initial:ThemeMode;preview:(value:ThemeMode)=>void;save:(value:ThemeMode)=>void|Promise<void>;onClose:()=>void}){
 const [original]=useState(initial);const [value,setValue]=useState(initial),[error,setError]=useState(''),[saving,setSaving]=useState(false);
 const cancel=()=>{if(!saving){preview(original);onClose();}};
 return <Modal title="外观设置" onClose={cancel}><form onSubmit={async event=>{event.preventDefault();setSaving(true);try{await save(value);onClose();}catch{setError('外观设置保存失败，请重试；取消可恢复原设置。');}finally{setSaving(false);}}}>
 <p className="organization-help">应用于此项目的所有工作区。选择后即时预览，保存后刷新保留。</p>
 <fieldset disabled={saving} className="theme-options"><legend>界面主题</legend>{([{mode:'light',label:'浅色',description:'明亮背景，适合日间阅读。'},{mode:'dark',label:'深色',description:'深色背景，减少夜间亮光。'},{mode:'system',label:'跟随系统',description:'随设备的浅色或深色设置变化。'}] as const).map(option=><label key={option.mode}><input type="radio" name="theme" value={option.mode} checked={value===option.mode} onChange={()=>{setValue(option.mode);preview(option.mode);setError('');}}/><span><strong>{option.label}</strong><small>{option.description}</small></span></label>)}</fieldset>
 <div className="theme-preview" aria-label="主题效果预览"><strong>记录今天的想法</strong><p>正文、侧栏和工具栏使用统一的外观。</p><span>文档标签</span><code>const idea = '继续写作';</code></div>
 {error&&<p role="alert" className="form-error">{error}</p>}<div className="dialog-actions"><button type="button" disabled={saving} onClick={cancel}>取消</button><button disabled={saving} className="primary">{error?'重试保存外观':'保存外观'}</button></div>
 </form></Modal>;
}
