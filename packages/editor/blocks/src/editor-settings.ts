export interface EditorSettings {width:'normal'|'wide'|'full';fontSize:14|16|18|20;lineHeight:1.5|1.8|2|2.2;selectionToolbar:boolean;mobileToolbar:boolean}
export const defaultEditorSettings:EditorSettings={width:'normal',fontSize:16,lineHeight:1.8,selectionToolbar:true,mobileToolbar:true};
export function readEditorSettings(value:unknown):EditorSettings {
 const raw=value&&typeof value==='object'?value as Partial<EditorSettings>:{};
 return {width:['normal','wide','full'].includes(raw.width??'')?raw.width!:defaultEditorSettings.width,fontSize:[14,16,18,20].includes(raw.fontSize??0)?raw.fontSize!:16,lineHeight:[1.5,1.8,2,2.2].includes(raw.lineHeight??0)?raw.lineHeight!:1.8,selectionToolbar:typeof raw.selectionToolbar==='boolean'?raw.selectionToolbar:true,mobileToolbar:typeof raw.mobileToolbar==='boolean'?raw.mobileToolbar:true};
}
