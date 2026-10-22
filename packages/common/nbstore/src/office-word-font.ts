/** Font size is stored in points; DOCX uses half-points. */
export function validWordFontSize(value:unknown):value is number{return typeof value==='number'&&Number.isInteger(value*2)&&value>=1&&value<=200;}
export function validWordColor(value:unknown):value is string{return value==='auto'||typeof value==='string'&&/^#[0-9A-Fa-f]{6}$/.test(value);}
/** Word's named highlight colours (w:highlight), with the colour used to draw them. */
export const WORD_HIGHLIGHTS:Record<string,string>={yellow:'#FFFF00',green:'#00FF00',cyan:'#00FFFF',magenta:'#FF00FF',blue:'#0000FF',red:'#FF0000',darkBlue:'#000080',darkCyan:'#008080',darkGreen:'#008000',darkMagenta:'#800080',darkRed:'#800000',darkYellow:'#808000',darkGray:'#808080',lightGray:'#C0C0C0',black:'#000000'};
export const validWordHighlight=(value:unknown):value is string=>typeof value==='string'&&Object.hasOwn(WORD_HIGHLIGHTS,value);
export const validWordVertAlign=(value:unknown):value is 'superscript'|'subscript'=>value==='superscript'||value==='subscript';
/** A font family name as stored in w:rFonts; no CSS syntax, so it can be quoted safely when rendered. */
export const validWordFontFamily=(value:unknown):value is string=>typeof value==='string'&&value.length>=1&&value.length<=64&&/^[\p{L}\p{N} ._-]+$/u.test(value)&&value.trim()===value;
export const WORD_FONT_FAMILIES=['宋体','黑体','楷体','仿宋','微软雅黑','等线','Calibri','Arial','Times New Roman','Cambria','Courier New'];
