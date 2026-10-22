import {DEFAULT_WORD_PAGE,validWordPage,type WordPageLayout} from '../../../../common/nbstore/src/office-word-page';
function attribute(xml:string,name:string){return new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`).exec(xml)?.[2];}
function attributes(xml:string,tag:string){return new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*)>`).exec(xml)?.[1];}
export function readWordPage(xml:string):WordPageLayout|undefined{
 const cleaned=xml.replace(/<w:sectPrChange\b[^>]*>[\s\S]*?<\/w:sectPrChange>/g,''),sections=[...cleaned.matchAll(/<w:sectPr(?=\s|\/|>)(?:[^>]*\/\s*>|[^>]*>([\s\S]*?)<\/w:sectPr>)/g)];
 if(sections.length>1||/<w:pPr(?:\s[^>]*)?>(?:(?!<\/w:pPr>)[\s\S])*<w:sectPr(?=\s|\/|>)/.test(cleaned))throw new Error('暂不支持多分节 Word 页面设置');
 if(!sections.length)return;
 const section=sections[0][1]??'',size=attributes(section,'pgSz'),margins=attributes(section,'pgMar'),page={...DEFAULT_WORD_PAGE};
 const number=(value:string)=>{if(!/^\d+$/.test(value))throw new Error('Word 页面尺寸或页边距无效');return Number(value);};
 for(const [xmlName,key]of [['w:w','width'],['w:h','height']] as const){const value=attribute(size??'',xmlName);if(value!==undefined)page[key]=number(value);}
 for(const key of ['top','right','bottom','left'] as const){const value=attribute(margins??'','w:'+key);if(value!==undefined)page[key]=number(value);}
 const orientation=attribute(size??'','w:orient');if(orientation&&(!['portrait','landscape'].includes(orientation)||(orientation==='landscape')!==(page.width>page.height)))throw new Error('Word 页面方向与尺寸不一致');
 const gutter=attribute(margins??'','w:gutter'),columns=attribute(attributes(section,'cols')??'','w:num');if(gutter&&gutter!=='0'||columns&&columns!=='1'||/<w:col(?=\s|\/|>)/.test(section))throw new Error('暂不支持装订线或多栏 Word 页面');
 if(!validWordPage(page))throw new Error('Word 页面尺寸或页边距超出范围');return page;
}
