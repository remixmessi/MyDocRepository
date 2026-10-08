import { documentSchema,markdownParser,readDocument,writeDocument } from '@mydoc/editor-blocks';
import { imageMime,MAX_ASSET_SIZE } from '@mydoc/nbstore';
import type { LocalWorkspace } from './workspace';
import { IMPORT_LIMITS,csvRows,notionName,notionId,resolveImportPath,readImportFiles,type ImportKind,type ImportFile } from './import-source';
type Doc=ReturnType<typeof readDocument>;
export interface ImportSource {kind:ImportKind;path:string;fingerprint:string}
export function validImportSource(value:unknown):value is ImportSource {if(!value||typeof value!=='object')return false;const s=value as ImportSource;return ['markdown','notion'].includes(s.kind)&&typeof s.path==='string'&&s.path.length<=1200&&/^[a-f\d]{64}$/.test(s.fingerprint);}
export interface ImportPage {path:string;title:string;folders:string[];node?:Doc;error?:string;fingerprint:string;warnings:string[]}
export interface ImportPlan {kind:ImportKind;pages:ImportPage[];files:Map<string,ImportFile>;warnings:string[]}
export interface ImportResult {path:string;title:string;status:'imported'|'skipped'|'failed';id?:string;message?:string;warnings:string[]}
export const importDuplicate=(ws:LocalWorkspace,page:ImportPage,kind:ImportKind)=>ws.list().find(doc=>doc.importSource?.kind===kind&&doc.importSource.path===page.path&&doc.importSource.fingerprint===page.fingerprint);
const digest=async(bytes:Uint8Array)=>[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(x=>x.toString(16).padStart(2,'0')).join('');
function csvDocument(text:string){const rows=csvRows(text);if(!rows.length)throw new Error('CSV 没有内容');const columns=Math.max(...rows.map(row=>row.length));const table=documentSchema.nodes.table.create(null,rows.map((row,index)=>documentSchema.nodes.table_row.create(null,Array.from({length:columns},(_,column)=>{const lines=(row[column]||'').split(/\r?\n/),content:Doc[]=[];lines.forEach((line,i)=>{if(i)content.push(documentSchema.nodes.hard_break.create());if(line)content.push(documentSchema.text(line));});return (index?documentSchema.nodes.table_cell:documentSchema.nodes.table_header).create(null,documentSchema.nodes.paragraph.create(null,content));}))));return documentSchema.nodes.doc.create(null,table);}
function sourceLinks(node:Doc){const links=new Set<string>();node.descendants(child=>{if(child.type.name==='image')links.add(child.attrs.src);for(const mark of child.marks)if(mark.type.name==='link')links.add(mark.attrs.href);});return [...links];}
export async function prepareImport(files:File[],kind:ImportKind,signal?:AbortSignal):Promise<ImportPlan>{
 let inputs=await readImportFiles(files,signal);const warnings:string[]=[];
 if(kind==='notion'){const roots=[...inputs.keys()].map(path=>path.split('/')[0]);if(roots.length&&roots.every(root=>root===roots[0])&&/^Export[-_ ]/i.test(roots[0]))inputs=new Map([...inputs.values()].map(file=>{const path=file.path.slice(roots[0].length+1);return [path,{...file,path}];}));}
 const pageFiles=[...inputs.values()].filter(file=>/\.(md|markdown|txt)$/i.test(file.path)||(kind==='notion'&&/\.csv$/i.test(file.path)));
 if(pageFiles.length>IMPORT_LIMITS.pages)throw new Error('每批最多导入 100 篇文档');if(!pageFiles.length)throw new Error('未找到可导入的 Markdown'+(kind==='notion'?' 或 CSV':' / 文本')+'文件');
 const hashes=new Map<string,string>();for(const file of inputs.values()){if(signal?.aborted)throw new DOMException('已取消','AbortError');hashes.set(file.path,await digest(file.bytes));}
 const pages:ImportPage[]=[];
 for(const file of pageFiles){if(signal?.aborted)throw new DOMException('已取消','AbortError');const parts=file.path.split('/'),stem=parts.pop()!.replace(/\.(md|markdown|txt|csv)$/i,'');const page:ImportPage={path:file.path,title:(kind==='notion'?notionName(stem):stem).slice(0,160)||'无标题文档',folders:parts.map(name=>kind==='notion'?notionName(name):name),fingerprint:'',warnings:[]};
  try{if(file.bytes.length>IMPORT_LIMITS.text)throw new Error('单篇文档不能超过 10 MB');let text:string;try{text=new TextDecoder('utf-8',{fatal:true}).decode(file.bytes).replace(/^\uFEFF/,'');}catch{throw new Error('文本编码不是 UTF-8，请转换编码后重试');}if(text.includes('\u0000'))throw new Error('文件不是可读取的文本');page.node=/\.csv$/i.test(file.path)?csvDocument(text):markdownParser.parse(text);
   const refs=sourceLinks(page.node).map(href=>{const target=resolveImportPath(file.path,href);return target?target+':'+(hashes.get(target)||'missing'):href;}).sort();page.fingerprint=await digest(new TextEncoder().encode(hashes.get(file.path)+'\n'+refs.join('\n')));
   for(const href of sourceLinks(page.node)){if(isExternal(href))continue;const path=resolveImportPath(file.path,href);if(!path||!inputs.has(path))page.warnings.push('未找到本地链接或资源：'+href);else if(!pageFiles.some(f=>f.path===path)&&inputs.get(path)!.bytes.length>MAX_ASSET_SIZE)page.warnings.push('资源超过 20 MB：'+href);}
   if(/<\/?(?:aside|details|summary|iframe)\b/i.test(text))page.warnings.push('HTML 标注或折叠内容按 Markdown 文本保留。');
  }catch(e){page.error=e instanceof Error?e.message:'文档无法解析';}pages.push(page);await new Promise(resolve=>setTimeout(resolve,0));
 }
 const unused=[...inputs.keys()].filter(path=>!pageFiles.some(file=>file.path===path));if(unused.length)warnings.push(`包内另有 ${unused.length} 个资源文件，仅保存正文引用的资源。`);
 return {kind,pages,files:inputs,warnings};
}
function findPage(plan:ImportPlan,from:string,href:string){const path=resolveImportPath(from,href);const exact=plan.pages.find(page=>page.path===path);if(exact)return exact;if(plan.kind==='notion'&&/^https?:\/\/(?:www\.)?notion\.(?:so|site)\//i.test(href)){try{const id=notionId(new URL(href).pathname.replaceAll('-',''));const found=plan.pages.filter(page=>notionId(page.path.replace(/\.[^.]+$/,''))===id);if(id&&found.length===1)return found[0];}catch{}}return undefined;}
const isExternal=(href:string)=>/^(?:[a-z][a-z\d+.-]*:|#|\/\/)/i.test(href);
export async function executeImport(ws:LocalWorkspace,plan:ImportPlan,selected:Set<string>,options:{folderId:string|null;hierarchy:boolean;duplicates:'skip'|'copy'},progress:(text:string)=>void,cancelled:()=>boolean=()=>false):Promise<ImportResult[]>{
 if(options.folderId&&!ws.folders.has(options.folderId))throw new Error('目标文件夹已删除');
 const results:ImportResult[]=[],ids=new Map<string,string>(),savedFiles=new Map<string,Promise<string>>(),folderIds=new Map<string,string|null>(),originals=new Map<string,Doc>();
 const pages=plan.pages.filter(page=>selected.has(page.path));if(options.duplicates==='skip')for(const page of plan.pages){const existing=importDuplicate(ws,page,plan.kind);if(existing)ids.set(page.path,existing.id);}let completed=0;
 const asset=async(path:string)=>{let promise=savedFiles.get(path);if(!promise){promise=(async()=>{const file=plan.files.get(path)!;const name=path.split('/').pop()!,mime=imageMime(file.bytes)||(/\.pdf$/i.test(name)?'application/pdf':'application/octet-stream');return (await ws.assets.save(new File([file.bytes],name,{type:mime}),false)).id;})();savedFiles.set(path,promise);}return promise;};
 async function rewrite(page:ImportPage,node:Doc,final:boolean,warnings:Set<string>):Promise<Doc>{
  const children:Doc[]=[];for(let i=0;i<node.childCount;i++){const child=node.child(i);
   if(child.type.name==='code_block'){children.push(child);continue;}
   if(child.type.name==='image'){const href=String(child.attrs.src);if(isExternal(href)){children.push(child);continue;}const path=resolveImportPath(page.path,href),file=path?plan.files.get(path):undefined;
    if(file){try{const id=await asset(path!);children.push(imageMime(file.bytes)?child.type.create({...child.attrs,src:'asset:'+id},null,child.marks):documentSchema.nodes.attachment.create({assetId:id,name:file.path.split('/').pop(),size:file.bytes.length,mime:'application/octet-stream'}));if(!imageMime(file.bytes))warnings.add('不支持的图片格式作为附件保存：'+href);}catch(e){warnings.add('资源未保存：'+href);children.push(documentSchema.text('图片未导入：'+href));}}
    else{warnings.add('图片资源缺失：'+href);children.push(documentSchema.text('图片缺失：'+href));}continue;}
   const link=child.marks.find(mark=>mark.type.name==='link');if(link){const href=String(link.attrs.href),target=findPage(plan,page.path,href),path=resolveImportPath(page.path,href),file=path?plan.files.get(path):undefined;let label=child.textContent;let end=i;while(end+1<node.childCount&&node.child(end+1).isText&&node.child(end+1).marks.some(mark=>mark.type.name==='link'&&mark.attrs.href===href)){end++;label+=node.child(end).textContent;}
    if(target&&final){i=end;const id=ids.get(target.path);if(id)children.push(documentSchema.nodes.doc_reference.create({docId:id,workspaceId:ws.id,label:label||target.title}));else{warnings.add('目标页面未导入：'+target.path);children.push(documentSchema.text(label+'（未导入页面）'));}continue;}
    if(!target&&!isExternal(href)){i=end;if(file){try{const id=await asset(path!);children.push(documentSchema.nodes.attachment.create({assetId:id,name:label||file.path.split('/').pop(),size:file.bytes.length,mime:'application/octet-stream'}));}catch{warnings.add('附件未保存：'+href);children.push(documentSchema.text(label+'（附件未导入）'));}}else{warnings.add('本地链接缺失：'+href);children.push(documentSchema.text(label+'（未找到本地文件）'));}continue;}
   }
   children.push(child.childCount?await rewrite(page,child,final,warnings):child);
  }
  return node.type.create(node.attrs,children,node.marks);
 }
 for(const page of pages){if(cancelled())break;progress(`正在导入 ${++completed}/${pages.length}：${page.title}`);if(page.error||!page.node){results.push({path:page.path,title:page.title,status:'failed',message:page.error,warnings:[]});continue;}
  const duplicate=options.duplicates==='skip'?importDuplicate(ws,page,plan.kind):undefined;if(duplicate){ids.set(page.path,duplicate.id);results.push({path:page.path,title:page.title,status:'skipped',id:duplicate.id,warnings:[]});continue;}
  const warnings=new Set(page.warnings);try{
   const rewritten=await rewrite(page,page.node,false,warnings);let folderId=options.folderId;
   if(options.hierarchy){let key='';for(const name of page.folders){key+='/'+name;if(folderIds.has(key)){folderId=folderIds.get(key)!;continue;}const existing=[...ws.folders.values()].find(folder=>folder.parentId===folderId&&folder.name===name);folderId=existing?.id||ws.createFolder(name.slice(0,160),folderId);folderIds.set(key,folderId);}}
   if(ws.persistence.state==='error')await ws.persistence.retry();const id=await ws.create(page.title,undefined,folderId,{content:rewritten});ids.set(page.path,id);originals.set(page.path,rewritten);results.push({path:page.path,title:page.title,status:'imported',id,warnings:[...warnings]});
  }catch(e){results.push({path:page.path,title:page.title,status:'failed',message:e instanceof Error?e.message:'保存失败',warnings:[...warnings]});}await new Promise(resolve=>setTimeout(resolve,0));
 }
 // Publish references only to documents that were actually saved (including skipped imports).
 for(const result of results){if(result.status!=='imported'||!result.id)continue;const page=pages.find(page=>page.path===result.path)!,item=await ws.openDocument(result.id),warnings=new Set(result.warnings),before=readDocument(item.store);
  try{progress('正在转换页面链接：'+page.title);const content=await rewrite(page,originals.get(page.path)!,true,warnings);writeDocument(item.store,content);await item.persistence.flush();const meta=ws.documents.get(result.id)!;
   if(![...warnings].some(w=>/缺失|未导入|未保存|未找到/.test(w)))ws.documents.set(result.id,{...meta,importSource:{kind:plan.kind,path:page.path,fingerprint:page.fingerprint}});await ws.persistence.flush();
  }catch(e){writeDocument(item.store,before);const meta=ws.documents.get(result.id);if(meta)ws.documents.set(result.id,{...meta,importSource:undefined});warnings.add('页面已创建，但链接转换或保存失败，请重试导入。');result.message=e instanceof Error?e.message:'保存失败';}result.warnings=[...warnings];
 }
 for(const page of pages)if(!results.some(result=>result.path===page.path))results.push({path:page.path,title:page.title,status:'failed',message:'已取消，未导入',warnings:[]});progress('导入完成');return results;
}
