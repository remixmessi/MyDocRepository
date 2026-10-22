export type WordAlignment='left'|'center'|'right'|'justify';
export const MAX_WORD_LEFT_INDENT=2880; // twips: 144 points, 2 inches.
export function validWordAlignment(value:unknown):value is WordAlignment{return typeof value==='string'&&['left','center','right','justify'].includes(value);}
export function validWordLeftIndent(value:unknown):value is number{return typeof value==='number'&&Number.isInteger(value)&&value>=0&&value<=MAX_WORD_LEFT_INDENT;}

export const validWordRightIndent=validWordLeftIndent;
export function validWordFirstLineIndent(value:unknown):value is number{return typeof value==='number'&&Number.isInteger(value)&&Math.abs(value)<=MAX_WORD_LEFT_INDENT;}
/** Line spacing as a multiple of single spacing (w:spacing w:line with lineRule="auto", 240 = 1.0). */
export function validWordLineSpacing(value:unknown):value is number{return typeof value==='number'&&Number.isFinite(value)&&value>=0.5&&value<=5&&Math.round(value*100)===value*100;}
/** Space before/after a paragraph in twips (w:spacing w:before/w:after), up to 72 pt. */
export function validWordParagraphSpace(value:unknown):value is number{return typeof value==='number'&&Number.isInteger(value)&&value>=0&&value<=1440;}
