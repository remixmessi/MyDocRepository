import katexCss from 'katex/dist/katex.min.css?raw';
const fontUrls=import.meta.glob('../../../../node_modules/katex/dist/fonts/*.woff2',{eager:true,query:'?url',import:'default'}) as Record<string,string>;
let cached:Promise<string>|undefined;
export function offlineKatexCss(){return cached??=(async()=>{
 const fonts=new Map<string,string>();await Promise.all(Object.entries(fontUrls).map(async([path,url])=>{const response=await fetch(url);if(!response.ok)throw new Error('公式字体读取失败，请重试');const bytes=new Uint8Array(await response.arrayBuffer());let binary='';for(let i=0;i<bytes.length;i+=8192)binary+=String.fromCharCode(...bytes.subarray(i,i+8192));fonts.set(path.split('/').at(-1)!,'data:font/woff2;base64,'+btoa(binary));}));
 return katexCss.replace(/src:([^;]+);/g,(_whole,src:string)=>{const filename=/([^/"'()]+\.woff2)/.exec(src)?.[1],data=filename&&fonts.get(filename);if(!data)throw new Error('公式字体不完整，请重试');return `src:url("${data}") format("woff2");`;});
 })().catch(error=>{cached=undefined;throw error;});}
