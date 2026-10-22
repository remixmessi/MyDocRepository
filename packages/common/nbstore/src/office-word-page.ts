export type WordPageLayout={width:number;height:number;top:number;right:number;bottom:number;left:number};
export const DEFAULT_WORD_PAGE:WordPageLayout={width:11906,height:16838,top:1440,right:1440,bottom:1440,left:1440};
export const WORD_PAGE_KEYS=['width','height','top','right','bottom','left'] as const;
export function validWordPage(value:unknown):value is WordPageLayout{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const page=value as Record<string,unknown>;
 if(Object.keys(page).length!==6||Object.keys(page).some(key=>!WORD_PAGE_KEYS.includes(key as typeof WORD_PAGE_KEYS[number]))||WORD_PAGE_KEYS.some(key=>typeof page[key]!=='number'||!Number.isInteger(page[key])))return false;
 const p=page as WordPageLayout;return p.width>=1440&&p.width<=31680&&p.height>=1440&&p.height<=31680&&[p.top,p.right,p.bottom,p.left].every(n=>n>=0&&n<=7200)&&p.left+p.right+360<=p.width&&p.top+p.bottom+360<=p.height;
}
export function wordPageKey(page:WordPageLayout){return WORD_PAGE_KEYS.map(key=>page[key]).join(':');}
