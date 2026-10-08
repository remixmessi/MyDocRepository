import {Readability} from '@mozilla/readability';
import DOMPurify from 'dompurify';
import {DOMParser as DocumentParser,Fragment,type Node} from 'prosemirror-model';
import {documentSchema,markdownParser,markdownSerializer} from '@mydoc/editor-blocks';
import {imageMime,type AssetRecord} from '@mydoc/nbstore';
import {clipUrl,type WebClipSource} from './clip-source';
import type {LocalWorkspace} from './workspace';
export interface ClipArticle { source:WebClipSource; markdown:string; images:number }
export function extractClip(html:string,requestedUrl:string,finalUrl=requestedUrl):ClipArticle {
 if(html.length>3*1024*1024)throw new Error('HTML 文件不能超过 3 MB');const url=clipUrl(finalUrl),requested=clipUrl(requestedUrl);
 const doc=new DOMParser().parseFromString(html,'text/html');if(doc.querySelectorAll('*').length>20000)throw new Error('网页结构过大，请上传精简的 HTML 文件');
 const originalTitle=doc.title;doc.querySelectorAll('script,style,noscript,iframe,object,embed,form,input,button,nav,footer,base').forEach(node=>node.remove());
 const base=doc.createElement('base');base.href=url;doc.head.append(base);
 doc.querySelectorAll('img').forEach(image=>{const src=image.getAttribute('data-src')||image.getAttribute('data-original')||image.getAttribute('src');if(src)try{image.setAttribute('src',new URL(src,url).href);}catch{image.remove();}});
 const article=new Readability(doc.cloneNode(true) as Document,{charThreshold:120,maxElemsToParse:20000}).parse();
 const content=article?.content||doc.querySelector('article,main')?.innerHTML||doc.body.innerHTML;
 const cleaned=DOMPurify.sanitize(content,{ALLOWED_TAGS:['p','br','div','section','article','h1','h2','h3','h4','h5','h6','strong','b','em','i','s','del','code','pre','blockquote','ul','ol','li','hr','a','img','table','thead','tbody','tr','th','td'],ALLOWED_ATTR:['href','src','alt','title','start','colspan','rowspan'],ALLOW_DATA_ATTR:false});
 const body=new DOMParser().parseFromString(cleaned,'text/html').body;
 for(const element of body.querySelectorAll('a,img')){const attr=element.tagName==='IMG'?'src':'href',raw=element.getAttribute(attr);try{if(!raw)throw 0;const target=new URL(raw,url);if(!['http:','https:'].includes(target.protocol)||target.username||target.password)throw 0;element.setAttribute(attr,target.href);}catch{if(attr==='src')element.remove();else element.removeAttribute(attr);}}
 const node=DocumentParser.fromSchema(documentSchema).parse(body);const markdown=markdownSerializer.serialize(node);
 if(node.textContent.trim().length<20)throw new Error('没有提取到足够正文；网页可能需要登录或动态加载，请上传保存的 HTML 文件');
 return {source:{url,requestedUrl:requested,title:(article?.title||originalTitle||new URL(url).hostname).trim().slice(0,500),clippedAt:Date.now()},markdown,images:body.querySelectorAll('img').length};
}
export async function clipFetch(url:string,kind:'page'|'image',signal:AbortSignal) {
 const controller=new AbortController();const abort=()=>controller.abort();signal.addEventListener('abort',abort,{once:true});if(signal.aborted)controller.abort();const timer=setTimeout(abort,20000);
 try{const response=await fetch('/api/clipper',{method:'POST',credentials:'omit',headers:{'Content-Type':'application/json'},body:JSON.stringify({url:clipUrl(url),kind}),signal:controller.signal});
 if(!response.ok){let message='网页读取失败；可重试或上传 HTML 文件';try{message=(await response.json()).error||message;}catch{}throw new Error(message);}
 if(kind==='page'&&!response.headers.get('content-type')?.includes('application/json'))throw new Error('当前部署没有网页读取接口，请上传 HTML 文件');const reader=response.body?.getReader();if(!reader)throw new Error('网页内容为空');const chunks:Uint8Array[]=[];let size=0;for(;;){const result=await reader.read();if(result.done)break;size+=result.value.length;if(size>20*1024*1024){await reader.cancel();throw new Error('网页或图片超过大小限制');}chunks.push(result.value);}const bytes=new Uint8Array(size);let offset=0;for(const chunk of chunks){bytes.set(chunk,offset);offset+=chunk.length;}return new Response(bytes,{status:response.status,headers:response.headers});
 }catch(e){if(controller.signal.aborted&&!signal.aborted)throw new Error('网页读取超时，请重试或上传 HTML 文件');throw e;}finally{clearTimeout(timer);signal.removeEventListener('abort',abort);}
}
export function clipDuplicate(workspace:LocalWorkspace,source:WebClipSource){const urls=new Set([source.url,source.requestedUrl]);return workspace.allDocuments().find(meta=>meta.webClip&&(urls.has(meta.webClip.url)||urls.has(meta.webClip.requestedUrl)));}
function mapImages(node:Node,assets:Map<string,AssetRecord>):Node {
 if(node.type.name==='image'){let key:string;try{key=clipUrl(node.attrs.src);}catch{key=node.attrs.src;}const asset=assets.get(key);if(asset)return node.type.create({...node.attrs,src:'asset:'+asset.id},null,node.marks);}
 if(node.isLeaf)return node;const children:Node[]=[];node.forEach(child=>children.push(mapImages(child,assets)));return node.copy(Fragment.fromArray(children));
}
export async function saveClip(workspace:LocalWorkspace,article:ClipArticle,options:{title:string;markdown:string;folderId:string|null;tagIds:string[];localImages:boolean;copy:boolean},signal:AbortSignal,progress:(message:string)=>void) {
 if(!options.title.trim()||options.title.length>500)throw new Error('请输入 1–500 字的标题');if(options.markdown.length>1024*1024)throw new Error('正文不能超过 1 MB');
 if(!options.copy&&clipDuplicate(workspace,article.source))throw new Error('相同来源已剪藏，请打开原文档或选择创建副本');
 let node=markdownParser.parse(options.markdown);if(!node.textContent.trim())throw new Error('正文不能为空');const downloaded=new Map<string,AssetRecord>(),stored:string[]=[];
 try{if(options.localImages){const urls=new Set<string>();node.descendants(n=>{if(n.type.name==='image'&&n.attrs.src)urls.add(clipUrl(n.attrs.src));});if(urls.size>30)throw new Error('本地图片最多 30 张，请减少图片或使用远程链接');let total=0;
 for(const url of urls){if(signal.aborted)throw new DOMException('已取消','AbortError');progress(`正在保存图片 ${downloaded.size+1}/${urls.size}…`);const response=await clipFetch(url,'image',signal),bytes=new Uint8Array(await response.arrayBuffer());const mime=imageMime(bytes);if(!mime||bytes.length>20*1024*1024)throw new Error('图片格式或大小无效，请使用远程链接');total+=bytes.length;if(total>50*1024*1024)throw new Error('图片总大小超过 50 MB，请使用远程链接');downloaded.set(url,{id:crypto.randomUUID(),name:(new URL(url).pathname.split('/').pop()||'网页图片').slice(0,200),mime,size:bytes.length,bytes});}
 for(const asset of downloaded.values()){if(signal.aborted)throw new DOMException('已取消','AbortError');await workspace.assets.put(asset);stored.push(asset.id);}node=mapImages(node,downloaded);}
 if(signal.aborted)throw new DOMException('已取消','AbortError');if(!options.copy&&clipDuplicate(workspace,article.source))throw new Error('相同来源已剪藏，请选择创建副本');
 const footer=Fragment.fromArray([documentSchema.nodes.horizontal_rule.create(),documentSchema.nodes.paragraph.create(null,[documentSchema.text('来源：'),documentSchema.text(new URL(article.source.url).hostname,[documentSchema.marks.link.create({href:article.source.url})]),documentSchema.text(' · 剪藏于 '+new Date(article.source.clippedAt).toLocaleString('zh-CN'))])]);node=node.copy(node.content.append(footer));progress('正在创建文档…');
 return await workspace.create(options.title.trim(),undefined,options.folderId,{content:node,properties:{webClip:article.source,tagIds:options.tagIds}});
 }catch(error){await Promise.allSettled(stored.map(id=>workspace.assets.remove(id)));throw error;}
}
