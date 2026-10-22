import os
os.chdir('packages/frontend/apps/web/server')
def edit(p,pairs):
    s=open(p).read()
    for a,b in pairs:
        assert s.count(a)==1,(p,a[:80]);s=s.replace(a,b)
    open(p,'w').write(s)
edit('office-word.ts',[
("import {validWordFontSize,validWordColor} from '../../../../common/nbstore/src/office-word-font';","import {validWordFontSize,validWordColor,validWordHighlight,validWordVertAlign,validWordFontFamily} from '../../../../common/nbstore/src/office-word-font';"),
("const MARKS=['bold','italic','underline'];","const MARKS=['bold','italic','underline','strike'];"),
(" for(const [tag,key]of [['b','bold'],['i','italic'],['u','underline']]){"," for(const [tag,key]of [['b','bold'],['i','italic'],['u','underline'],['strike','strike']]){"),
(""" return attributes;
}
export function encodeDocxWord""",r""" const vert=/<w:vertAlign(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:vertAlign>)/.exec(properties);
 if(vert){const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(vert[1])?.[2];if(validWordVertAlign(value))attributes.vertAlign=value;else if(value!=='baseline')throw new Error('Word 上下标设置无效');}
 const highlight=/<w:highlight(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:highlight>)/.exec(properties);
 if(highlight){const value=/\bw:val\s*=\s*(["'])(.*?)\1/.exec(highlight[1])?.[2];if(validWordHighlight(value))attributes.highlight=value;else if(value!=='none')throw new Error('暂不支持该 Word 突出显示颜色');}
 // Explicit font names only; theme font references are left to the document default.
 const fonts=/<w:rFonts(?=\s|\/|>)([^>]*?)(?:\/>|>[\s\S]*?<\/w:rFonts>)/.exec(properties);
 if(fonts){const pick=(name:string)=>{const value=new RegExp(`\\bw:${name}\\s*=\\s*(["'])(.*?)\\1`).exec(fonts[1])?.[2];return value===undefined?undefined:decodeWordXml(value);},family=pick('eastAsia')??pick('ascii')??pick('hAnsi');if(family!==undefined&&validWordFontFamily(family))attributes.fontFamily=family;}
 return attributes;
}
export function encodeDocxWord"""),
("key==='list'?!validWordList(attribute):","key==='list'?!validWordList(attribute):key==='lineSpacing'?!validWordLineSpacing(attribute):key==='spaceBefore'||key==='spaceAfter'?!validWordParagraphSpace(attribute):"),
("key==='fontSize'?!validWordFontSize(attribute):key==='color'?!validWordColor(attribute):","key==='fontSize'?!validWordFontSize(attribute):key==='color'?!validWordColor(attribute):key==='highlight'?!validWordHighlight(attribute):key==='vertAlign'?!validWordVertAlign(attribute):key==='fontFamily'?!validWordFontFamily(attribute):"),
("import {validWordAlignment,validWordLeftIndent,validWordRightIndent,validWordFirstLineIndent} from '../../../../common/nbstore/src/office-word-layout';","import {validWordAlignment,validWordLeftIndent,validWordRightIndent,validWordFirstLineIndent,validWordLineSpacing,validWordParagraphSpace} from '../../../../common/nbstore/src/office-word-layout';"),
])
edit('office-word-layout.ts',[
("import {validWordAlignment,validWordLeftIndent,validWordFirstLineIndent,type WordAlignment} from '../../../../common/nbstore/src/office-word-layout';","import {validWordAlignment,validWordLeftIndent,validWordFirstLineIndent,validWordLineSpacing,validWordParagraphSpace,type WordAlignment} from '../../../../common/nbstore/src/office-word-layout';"),
("type Layout={pageBreakBefore?:boolean;align?:WordAlignment;leftIndent?:number;rightIndent?:number;firstLineIndent?:number};","type Layout={pageBreakBefore?:boolean;align?:WordAlignment;leftIndent?:number;rightIndent?:number;firstLineIndent?:number;lineSpacing?:number;spaceBefore?:number;spaceAfter?:number};"),
("""
 return layout;
}
export function wordLayoutReader""",r"""
 // Spacing: auto line rule is a multiple of 240; exact/at-least heights and auto spacing are not modelled and stay at the default.
 const spacing=property(xml,'spacing');
 if(spacing!==undefined){
  const number=(key:string)=>{const value=attribute(spacing,key);return value!==undefined&&/^\d+$/.test(value)?Number(value):undefined;};
  const line=number('w:line'),rule=attribute(spacing,'w:lineRule')??'auto';if(line!==undefined&&rule==='auto'){const multiple=Math.round(line/240*100)/100;if(validWordLineSpacing(multiple))layout.lineSpacing=multiple;}
  for(const [key,field,auto]of [['w:before','spaceBefore','w:beforeAutospacing'],['w:after','spaceAfter','w:afterAutospacing']] as const){const value=number(key);if(value!==undefined&&!['1','true','on'].includes(attribute(spacing,auto)??''))layout[field]=Math.min(value,1440);}
 }
 return layout;
}
export function wordLayoutReader"""),
])
edit('office-word-run-style.ts',[
("export type WordRunMarks={bold?:boolean;italic?:boolean;underline?:boolean;fontSize?:number;color?:string};","export type WordRunMarks={bold?:boolean;italic?:boolean;underline?:boolean;strike?:boolean;fontSize?:number;color?:string};"),
("const TOGGLES=[['b','bold'],['i','italic']] as const;","const TOGGLES=[['b','bold'],['i','italic'],['strike','strike']] as const;"),
("  for(const key of ['bold','italic','underline'] as const)if(inherited[key]","  for(const key of ['bold','italic','underline','strike'] as const)if(inherited[key]"),
])
print('done')
