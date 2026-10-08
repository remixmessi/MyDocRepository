import katex from 'katex';
import 'katex/dist/katex.min.css';
import DOMPurify from 'dompurify';
import {expressionLimit} from './expression-format';
import {isDarkTheme} from './visual-theme';
let serial=0;
const cache=new Map<string,string>();let cachedSize=0;
let engine:Promise<typeof import('mermaid').default>|undefined;
function mermaidEngine(){return engine??=(import('mermaid').then(({default:api})=>{api.initialize({startOnLoad:false,securityLevel:'strict',suppressErrorRendering:true,maxTextSize:20000,maxEdges:500,htmlLabels:false,flowchart:{htmlLabels:false},secure:['securityLevel','maxTextSize','maxEdges','startOnLoad','suppressErrorRendering','htmlLabels'],theme:'neutral'});return api;}).catch(e=>{engine=undefined;throw e;}));}

let renderQueue:Promise<unknown>=Promise.resolve();
function renderDiagram(source:string,dark:boolean){const job=renderQueue.then(async()=>{const api=await mermaidEngine();api.initialize({startOnLoad:false,securityLevel:'strict',suppressErrorRendering:true,maxTextSize:20000,maxEdges:500,htmlLabels:false,flowchart:{htmlLabels:false},secure:['securityLevel','maxTextSize','maxEdges','startOnLoad','suppressErrorRendering','htmlLabels'],theme:dark?'dark':'neutral'});return api.render(`mydoc-diagram-${++serial}`,source);});renderQueue=job.catch(()=>{});return job;}
export function paintExpression(host:HTMLElement,type:string,source:string){let cleanup=paintExpressionFrame(host,type,source,isDarkTheme()&&!host.closest('.board-note'));const refresh=()=>{if(type!=='mermaid')return;cleanup();cleanup=paintExpressionFrame(host,type,source,isDarkTheme()&&!host.closest('.board-note'));};window.addEventListener('mydoc-theme-change',refresh);return()=>{cleanup();window.removeEventListener('mydoc-theme-change',refresh);};}

function paintExpressionFrame(host:HTMLElement,type:string,source:string,dark=false){let alive=true;host.replaceChildren();host.className='expression-preview';const message=(text:string,role:string)=>{const node=document.createElement('span');node.textContent=text;node.setAttribute('role',role);host.replaceChildren(node);};
 if(source.length>expressionLimit(type)){message('源码超过长度限制','alert');return()=>{alive=false;};}
 if(!source.trim()){message(type==='mermaid'?'暂无图表，请编辑 Mermaid 源码':'暂无公式，请编辑 LaTeX 源码','status');return()=>{alive=false;};}
 if(type!=='mermaid'){try{katex.render(source,host,{displayMode:type==='math_block',throwOnError:true,trust:false,strict:'ignore',maxExpand:1000,maxSize:20});}catch{message('公式语法有误，请编辑源码修正。','alert');}return()=>{alive=false;};}
 if(/%%\{|^\s*---\s*\n/.test(source)){message('图表不支持配置指令，请保留图表源码。','alert');return()=>{alive=false;};}
 const install=(svg:string)=>{const prefix=`mydoc-preview-${++serial}-`,ids=new Map<string,string>();svg=svg.replace(/\bid="([^"]+)"/g,(_match,id:string)=>{ids.set(id,prefix+id);return `id="${prefix+id}"`;});for(const [id,next]of [...ids].sort((a,b)=>b[0].length-a[0].length))svg=svg.replaceAll('#'+id,'#'+next);svg=svg.replace(/\baria-(labelledby|describedby)="([^"]*)"/g,(_match,name:string,value:string)=>`aria-${name}="${value.split(' ').map(id=>ids.get(id)??id).join(' ')}"`);host.innerHTML=svg;const drawing=host.querySelector('svg');if(drawing){drawing.setAttribute('role','img');drawing.setAttribute('aria-label','Mermaid 图表');const width=Number(drawing.getAttribute('viewBox')?.split(/[\s,]+/)[2]);if(Number.isFinite(width)&&width>0)drawing.style.width=width+'px';drawing.style.maxWidth='100%';}};
 const cacheKey=(dark?'dark:':'light:')+source;
 if(cache.has(cacheKey)){install(cache.get(cacheKey)!);return()=>{alive=false;};}
 message('正在绘制图表…','status');
 void renderDiagram(source,dark).then(result=>{if(!alive)return;const svg=DOMPurify.sanitize(result.svg,{USE_PROFILES:{svg:true,svgFilters:true},FORBID_TAGS:['foreignObject'],FORBID_ATTR:['href','xlink:href']});install(svg);if(svg.length<2000000){while(cache.size>=12||cachedSize+svg.length>2000000){const first=cache.keys().next().value;if(first===undefined)break;cachedSize-=cache.get(first)!.length;cache.delete(first);}cache.set(cacheKey,svg);cachedSize+=svg.length;}}).catch(()=>{if(alive)message('图表语法有误或渲染失败，请修改源码后重试。','alert');});
 return()=>{alive=false;};
}
export function hydrateExpressions(host:HTMLElement){const cleanups:(()=>void)[]=[];host.querySelectorAll<HTMLElement>('[data-expression]').forEach(el=>{cleanups.push(paintExpression(el,el.dataset.expression!,el.dataset.source||''));});return()=>cleanups.forEach(clean=>clean());}

/** Resolve an expression to static markup; exported documents do not require JavaScript. */
export async function expressionSnapshot(type:string,source:string):Promise<string>{
 if(source.length>expressionLimit(type)||!source.trim())throw new Error('表达内容为空或超过长度限制');
 if(type!=='mermaid')return katex.renderToString(source,{displayMode:type==='math_block',throwOnError:true,trust:false,strict:'ignore',maxExpand:1000,maxSize:20});
 if(/%%\{|^\s*---\s*\n/.test(source))throw new Error('图表不支持配置指令');
 const raw=cache.get('light:'+source)??DOMPurify.sanitize((await renderDiagram(source,false)).svg,{USE_PROFILES:{svg:true,svgFilters:true},FORBID_TAGS:['foreignObject'],FORBID_ATTR:['href','xlink:href']});
 const prefix=`mydoc-export-${++serial}-`,ids=new Map<string,string>();let svg=raw.replace(/\bid="([^"]+)"/g,(_match,id:string)=>{ids.set(id,prefix+id);return `id="${prefix+id}"`;});
 for(const[id,next]of [...ids].sort((a,b)=>b[0].length-a[0].length))svg=svg.replaceAll('#'+id,'#'+next);
 return svg.replace(/\baria-(labelledby|describedby)="([^"]*)"/g,(_match,name:string,value:string)=>`aria-${name}="${value.split(' ').map(id=>ids.get(id)??id).join(' ')}"`);
}
