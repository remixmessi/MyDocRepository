import {validWordLink} from '../../../common/nbstore/src/office-word-link';
import {WORD_HIGHLIGHTS,validWordHighlight,validWordFontFamily,validWordVertAlign} from '../../../common/nbstore/src/office-word-font';
export type WordPreviewPart={insert:unknown;attributes?:Record<string,boolean|string|number>};
export function WordPreviewRun({part}:{part:WordPreviewPart}){
  const marks=part.attributes??{};
  const vert=validWordVertAlign(marks.vertAlign)?marks.vertAlign:undefined;
  const style={fontSize:vert?(typeof marks.fontSize==='number'?`${marks.fontSize*0.65}pt`:'0.65em'):typeof marks.fontSize==='number'?`${marks.fontSize}pt`:undefined,verticalAlign:vert?(vert==='superscript'?'super':'sub'):undefined,color:typeof marks.color==='string'?(marks.color==='auto'?'var(--text)':marks.color):undefined,backgroundColor:validWordHighlight(marks.highlight)?WORD_HIGHLIGHTS[marks.highlight]:undefined,fontFamily:validWordFontFamily(marks.fontFamily)?`"${marks.fontFamily}"`:undefined,fontWeight:marks.bold===undefined?undefined:marks.bold?'bold':'normal',fontStyle:marks.italic===undefined?undefined:marks.italic?'italic':'normal',textDecoration:marks.underline===undefined&&marks.strike===undefined?undefined:[marks.underline?'underline':'',marks.strike?'line-through':''].filter(Boolean).join(' ')||'none'};
  return validWordLink(marks.href)?<a href={marks.href} target="_blank" rel="noopener noreferrer" title={marks.href} style={style}><span>{String(part.insert??'')}</span></a>:<span style={style}>{String(part.insert??'')}</span>;
}

