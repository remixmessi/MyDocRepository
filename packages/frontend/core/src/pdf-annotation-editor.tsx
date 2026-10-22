import {officeAnchorTarget} from '../../../common/nbstore/src/office-comment-anchor';
import type {PdfFormSpec} from '../../../common/nbstore/src/pdf-forms';
import {PdfFormPanel} from './pdf-form-panel';
import {PdfNativeAnnotations} from './pdf-native-annotations';
import {useOfficePresence} from './use-office-presence';
import {OfficePresenceList} from './office-presence-list';
import {OfficeComments} from './office-comments-ui';
import {OfficeActivityButton,useFollowPeer,requestOfficeThread} from './office-collaboration-ui';
import {validOfficeComment,type OfficeCommentThread} from '../../../common/nbstore/src/office-comments';
import type {OfficePeer} from '../../../common/nbstore/src/office-presence';
import {useOfficeDocument} from './use-office-document';
import {useEffect,useRef,useState} from 'react';
import * as Y from 'yjs';
import type {LocalWorkspace,OfficePdfDocument} from './workspace';
import {CLOUD_ORIGIN,cloudJson,cloudRequest,officeDocumentGeneration} from '../../../common/nbstore/src/cloud';
import {pdfAnnotationId,pdfAnnotationRevision,validPdfAnnotation,type PdfAnnotation,type PdfOperation} from '../../../common/nbstore/src/office-pdf';
import {PdfPageView,type PdfZoom} from './pdf-page-view';
import {PdfThumbnails} from './pdf-thumbnails';
import {usePdfDocument,searchPdf} from './use-pdf-document';
import {OfficeTitleBar,OfficeRibbon,RibbonGroup,RibbonButton,OfficeStatusBar} from './office-ribbon';
import {ChevronLeft,ChevronRight,ZoomIn,ZoomOut,RotateCw,TextCursor,StickyNote,Highlighter,PenLine,Download,FileDown,PanelLeft,Search,MoveHorizontal,Maximize} from 'lucide-react';
import {downloadOfficeBytes} from './office-download';
export function PdfAnnotationEditor({item,workspace,name}:{item:OfficePdfDocument;workspace:LocalWorkspace;name:string}){
 useOfficeDocument(item.doc,workspace);
 const [commentTarget,setCommentTarget]=useState<string>();
 const [formSpecs,setFormSpecs]=useState<PdfFormSpec[]>([]),[visibleFields,setVisibleFields]=useState<string[]>([]);
 const draftKey='pdf-draft:'+item.id,stored=useRef(workspace.uiPreferences.get(draftKey) as {annotation?:PdfAnnotation;editing?:{id:string;base:string};pending?:PdfOperation}|undefined),restored=stored.current?.annotation&&validPdfAnnotation(stored.current.annotation)?stored.current:undefined;
 const [page,setPage]=useState(String(restored?.annotation?.page??1)),[content,setContent]=useState(restored?.annotation?.content??''),[type,setType]=useState<PdfAnnotation['type']>(restored?.annotation?.type??'comment'),[geometry,setGeometry]=useState<Pick<PdfAnnotation,'x'|'y'|'width'|'height'|'points'>>(restored?.annotation?{x:restored.annotation.x,y:restored.annotation.y,...(restored.annotation.width?{width:restored.annotation.width}:{}),...(restored.annotation.height?{height:restored.annotation.height}:{}),...(restored.annotation.points?{points:restored.annotation.points}:{})}:{x:0,y:0}),[editing,setEditing]=useState<{id:string;base:string}|undefined>(restored?.editing),[busy,setBusy]=useState(false),[error,setError]=useState(restored?'已恢复本账户的批注草稿，请核对最新内容再提交。':'');const pending=useRef<PdfOperation|undefined>(restored?.pending);

 if(pending.current&&!Object.hasOwn(pending.current,'officeGeneration')&&officeDocumentGeneration(workspace.id,item.id)?.startsWith('epoch-'))(pending.current as unknown as Record<string,unknown>).officeGeneration='legacy-unknown';
 const readOnly=workspace.documentReadOnly(item.id),annotations=item.annotations.toArray(),annotation:PdfAnnotation={page:Number(page),type,content,...geometry};
 const document=usePdfDocument(workspace.id,item.id),pages=document.pdf?.numPages??0,current=Number.isSafeInteger(Number(page))&&Number(page)>0?Number(page):1;
 const [tab,setTab]=useState('开始'),[zoom,setZoom]=useState<PdfZoom>('fit'),[scale,setScale]=useState(1),[rotation,setRotation]=useState(0),[selectText,setSelectText]=useState(false),[side,setSide]=useState<'thumbnails'|'search'|undefined>(()=>typeof matchMedia==='function'&&matchMedia('(max-width: 900px)').matches?undefined:'thumbnails');
 const [query,setQuery]=useState(''),[hits,setHits]=useState<{page:number;excerpt:string}[]>(),[searching,setSearching]=useState(false),[highlight,setHighlight]=useState<string>();
 const goto=(value:number)=>{if(!pages)return;setPage(String(Math.max(1,Math.min(pages,value))));setGeometry({x:0,y:0});};
 const step=(direction:1|-1)=>{const levels=[0.5,0.75,1,1.25,1.5,2,3],now=typeof zoom==='number'?zoom:scale,next=direction>0?levels.find(level=>level>now+0.01):[...levels].reverse().find(level=>level<now-0.01);if(next)setZoom(next);};
 const tool=(value:PdfAnnotation['type'])=>{setType(value);setGeometry({x:0,y:0});setSelectText(false);};
 const counts=new Map<number,number>();for(const value of annotations)counts.set(Number(value.get('page')),(counts.get(Number(value.get('page')))??0)+1);
 // Unresolved discussions per page (thumbnail badge) and per annotation (list button).
 const discussions=new Map<number,number>(),annotationThreads=new Map<string,string[]>();if(item.doc.share.has('officeComments'))for(const thread of item.doc.getMap<OfficeCommentThread>('officeComments').values()){if(!validOfficeComment(thread)||thread.resolved||thread.anchor?.kind!=='pdf')continue;discussions.set(thread.anchor.page,(discussions.get(thread.anchor.page)??0)+1);annotationThreads.set(thread.anchor.annotation,[...(annotationThreads.get(thread.anchor.annotation)??[]),thread.id]);}
 const search=async()=>{if(!document.pdf||!query.trim())return;setSearching(true);setSide('search');const signal={cancelled:false};try{const result=await searchPdf(document.pdf,query,signal);setHits(result);setHighlight(query);if(result[0])goto(result[0].page);}finally{setSearching(false);}};
 const peers=useOfficePresence(item.doc,'pdf',Number.isSafeInteger(Number(page))&&Number(page)>=1?{kind:'pdf',page:Number(page)}:undefined,!!workspace.sync);
 const locatePeer=(peer:OfficePeer)=>{if(peer.focus?.kind==='pdf')goto(peer.focus.page);},shell=useRef<HTMLElement>(null),follow=useFollowPeer(peers,locatePeer,shell);
 useEffect(()=>{if(busy)return;if(content||editing||pending.current||type==='highlight'&&geometry.width||type==='ink'&&geometry.points){workspace.uiPreferences.set(draftKey,JSON.parse(JSON.stringify({annotation,editing,pending:pending.current})));}else workspace.uiPreferences.delete(draftKey);},[page,content,type,geometry,editing,busy,error,workspace,draftKey]);
 const hasDraft=!!(content||editing||pending.current||type==='highlight'&&geometry.width||type==='ink'&&geometry.points);
 const clearDraft=()=>{setEditing(undefined);setContent('');setGeometry({x:0,y:0});pending.current=undefined;setError('');workspace.uiPreferences.delete(draftKey);};
 const apply=(update:string)=>Y.applyUpdate(item.doc,Uint8Array.from(atob(update),char=>char.charCodeAt(0)),CLOUD_ORIGIN);
 const submit=async(operation:Omit<PdfOperation,'operationId'>)=>{
  if(busy||readOnly)return;setBusy(true);setError('');const old=pending.current;if(!old||JSON.stringify({...old,operationId:undefined,officeGeneration:undefined})!==JSON.stringify({...operation,officeGeneration:undefined}))pending.current={...operation,operationId:crypto.randomUUID()};
  try{const response=await cloudJson<{update:string}>('/office-pdf-operation?'+new URLSearchParams({workspace:workspace.id,id:item.id}),'POST',pending.current);apply(response.update);pending.current=undefined;setEditing(undefined);setContent('');setGeometry({x:0,y:0});workspace.uiPreferences.delete(draftKey);}catch(error){setError(error instanceof Error?error.message:'批注提交失败，草稿已保留');}finally{setBusy(false);}
 };
 const select=(value:Y.Map<unknown>)=>{setCommentTarget(pdfAnnotationId(value));setEditing({id:pdfAnnotationId(value),base:pdfAnnotationRevision(value)});setPage(String(value.get('page')));setType(value.get('type') as PdfAnnotation['type']);setContent(String(value.get('content')));const a=value.toJSON() as unknown as PdfAnnotation;setGeometry({x:a.x,y:a.y,...(a.width?{width:a.width}:{}),...(a.height?{height:a.height}:{}),...(a.points?{points:a.points}:{})});pending.current=undefined;setError('');};
 const latest=async()=>{setBusy(true);try{const response=await cloudRequest('/document?'+new URLSearchParams({workspace:workspace.id,id:item.id}));Y.applyUpdate(item.doc,new Uint8Array(await response.arrayBuffer()),CLOUD_ORIGIN);const current=item.annotations.toArray().find(a=>pdfAnnotationId(a)===editing?.id);if(current){setEditing({id:pdfAnnotationId(current),base:pdfAnnotationRevision(current)});setError('已载入最新修订；当前输入仍保留，请核对后提交。');}else{setEditing(undefined);setError('原批注已删除，输入保留，可另存为新批注。');}pending.current=undefined;}catch(error){setError(error instanceof Error?error.message:'读取最新内容失败');}finally{setBusy(false);}};
 const exportAnnotations=()=>downloadOfficeBytes(name.replace(/\.[^.]+$/,'')+'.annotations.json',new TextEncoder().encode(JSON.stringify({format:'mydoc-pdf-annotations',version:1,annotations:annotations.map(a=>a.toJSON())},null,2)),'application/json');
 const exportPdf=async(flatten=false)=>{setBusy(true);setError('');try{await workspace.prepareOfficeBackup(item.id);const response=await cloudRequest('/office-pdf-export?'+new URLSearchParams({workspace:workspace.id,id:item.id,...(flatten?{flatten:'1'}:{})}));downloadOfficeBytes(name.replace(/\.pdf$/i,'')+(flatten?'-扁平化副本.pdf':'-批注副本.pdf'),new Uint8Array(await response.arrayBuffer()),'application/pdf');}catch(error){setError(error instanceof Error?error.message:'PDF 导出失败');}finally{setBusy(false);}};
 return <section ref={shell} className="office-pdf-annotations office-document-shell" aria-label={name}>
  <OfficeTitleBar kind="pdf" name={name} status={(readOnly?'只读':workspace.sync?'多人协同批注':'本地批注')+(pages?` · 共 ${pages} 页`:'')}>
   <OfficeActivityButton doc={item.doc} workspace={workspace} id={item.id}/><OfficeComments doc={item.doc} workspace={workspace} id={item.id} captureAnchor={()=>{const selected=annotations.find(value=>pdfAnnotationId(value)===commentTarget);if(selected)return {kind:'pdf',annotation:pdfAnnotationId(selected),page:Number(selected.get('page'))};}} onLocate={anchor=>{const target=officeAnchorTarget(item.doc,anchor);if(target?.kind==='pdf'){setPage(String(target.page));setCommentTarget(pdfAnnotationId(target.annotation));}}}/>
  </OfficeTitleBar>
  <OfficeRibbon tabs={['开始','批注','视图']} tab={tab} onTab={setTab} label="PDF 功能区">
   {tab==='开始'&&<>
    <RibbonGroup label="页面"><RibbonButton icon={ChevronLeft} label="上一页" disabled={!pages||current<=1} onClick={()=>goto(current-1)}/><RibbonButton icon={ChevronRight} label="下一页" disabled={!pages||current>=pages} onClick={()=>goto(current+1)}/></RibbonGroup>
    <RibbonGroup label="显示"><RibbonButton icon={ZoomOut} label="缩小" compact onClick={()=>step(-1)}/><select aria-label="PDF 缩放" value={String(zoom)} onChange={event=>setZoom(event.target.value==='fit'||event.target.value==='page'?event.target.value:Number(event.target.value))}><option value="fit">适合宽度</option><option value="page">适合页面</option>{typeof zoom==='number'&&![0.5,1,1.5,2].includes(zoom)&&<option value={zoom}>{Math.round(zoom*100)}%</option>}{[0.5,1,1.5,2].map(n=><option key={n} value={n}>{n*100}%</option>)}</select><RibbonButton icon={ZoomIn} label="放大" compact onClick={()=>step(1)}/><RibbonButton icon={RotateCw} label="旋转 90°" onClick={()=>setRotation(value=>(value+90)%360)}/></RibbonGroup>
    <RibbonGroup label="工具"><RibbonButton icon={TextCursor} label="选择文字" pressed={selectText} onClick={()=>setSelectText(value=>!value)}/>{(['comment','highlight','ink'] as const).map(value=><RibbonButton key={value} icon={{comment:StickyNote,highlight:Highlighter,ink:PenLine}[value]} label={{comment:'附注',highlight:'高亮',ink:'手绘'}[value]} pressed={!selectText&&type===value} disabled={readOnly||busy} onClick={()=>tool(value)}/>)}</RibbonGroup>
    <RibbonGroup label="导出"><RibbonButton icon={Download} label="导出带批注 PDF" disabled={busy} onClick={()=>void exportPdf()}/><RibbonButton icon={FileDown} label="导出扁平化 PDF" disabled={busy} onClick={()=>void exportPdf(true)}/></RibbonGroup>
   </>}
   {tab==='批注'&&<>
    <RibbonGroup label="批注工具">{(['comment','highlight','ink'] as const).map(value=><RibbonButton key={value} icon={{comment:StickyNote,highlight:Highlighter,ink:PenLine}[value]} label={{comment:'附注',highlight:'高亮',ink:'手绘'}[value]} large pressed={!selectText&&type===value} disabled={readOnly||busy} onClick={()=>tool(value)}/>)}</RibbonGroup>
    <RibbonGroup label="批注数据"><RibbonButton icon={Download} label="导出批注" disabled={busy} onClick={exportAnnotations}/></RibbonGroup>
   </>}
   {tab==='视图'&&<>
    <RibbonGroup label="侧边栏"><RibbonButton icon={PanelLeft} label="缩略图" pressed={side==='thumbnails'} onClick={()=>setSide(value=>value==='thumbnails'?undefined:'thumbnails')}/><RibbonButton icon={Search} label="查找" pressed={side==='search'} onClick={()=>setSide(value=>value==='search'?undefined:'search')}/></RibbonGroup>
    <RibbonGroup label="页面显示"><RibbonButton icon={MoveHorizontal} label="适合宽度" pressed={zoom==='fit'} onClick={()=>setZoom('fit')}/><RibbonButton icon={Maximize} label="适合页面" pressed={zoom==='page'} onClick={()=>setZoom('page')}/><RibbonButton label="实际大小" pressed={zoom===1} onClick={()=>setZoom(1)}/></RibbonGroup>
   </>}
  </OfficeRibbon>
  <div className="pdf-workspace" data-side={side??'none'}>
   {side&&<aside className="pdf-side" aria-label={side==='thumbnails'?'页面缩略图面板':'查找面板'}>
    <div className="pdf-side-tabs" role="tablist" aria-label="侧边栏"><button role="tab" aria-selected={side==='thumbnails'} onClick={()=>setSide('thumbnails')}>缩略图</button><button role="tab" aria-selected={side==='search'} onClick={()=>setSide('search')}>查找</button></div>
    {side==='thumbnails'&&(document.pdf?<PdfThumbnails pdf={document.pdf} page={current} counts={counts} discussions={discussions} onPage={goto}/>:<p className="pdf-side-empty">{document.error||'正在载入页面…'}</p>)}
    {side==='search'&&<div className="pdf-search"><form role="search" onSubmit={event=>{event.preventDefault();void search();}}><input aria-label="查找 PDF 文字" placeholder="查找文字" value={query} onChange={event=>setQuery(event.target.value)}/><button disabled={!document.pdf||searching||!query.trim()}>{searching?'查找中…':'查找'}</button></form>
     {hits&&<p role="status">{hits.length?`共 ${hits.length} 处${hits.length>=200?'（仅显示前 200 处）':''}`:'未找到匹配文字'}</p>}<ol>{hits?.map((hit,i)=><li key={i}><button onClick={()=>{goto(hit.page);setHighlight(query);}}><strong>第 {hit.page} 页</strong><span>{hit.excerpt}</span></button></li>)}</ol></div>}
   </aside>}
   <div className="pdf-center" tabIndex={-1} onKeyDown={event=>{if(event.target!==event.currentTarget)return;if(event.key==='PageDown'||event.key==='ArrowRight'){event.preventDefault();goto(current+1);}if(event.key==='PageUp'||event.key==='ArrowLeft'){event.preventDefault();goto(current-1);}}}>
    {document.error&&<p role="alert">{document.error}<button onClick={document.reload}>重试阅读</button></p>}{document.loading&&<p role="status">正在显示 PDF…</p>}
    <PdfPageView pdf={document.pdf} page={current} zoom={zoom} rotation={rotation} selectText={selectText} highlight={highlight} onScale={setScale} formSpecs={formSpecs} formItem={item} formWorkspace={workspace} onFields={setVisibleFields} type={type} annotations={annotations.map(a=>a.toJSON() as unknown as PdfAnnotation)} onLocate={setGeometry} readOnly={readOnly||busy||selectText}/>
   </div>
   <aside className="pdf-inspector" aria-label="批注属性">
    <h3>{editing?'修改批注':'新建批注'}</h3><p className="pdf-hint">原件不会被改写，批注以标准注释导出。{selectText?'选择文字模式：可复制原文。':type==='comment'?'在页面上单击放置附注。':type==='highlight'?'在页面上拖动框选高亮区域。':'在页面上按住拖动绘制笔迹。'}{readOnly?' 当前只读。':''}</p>
    <form onSubmit={event=>{event.preventDefault();void submit({type:editing?'update':'create',annotationId:editing?.id??(pending.current?.type==='create'?pending.current.annotationId:crypto.randomUUID()),base:editing?.base??'',annotation});}}>
   <label>批注类型<select aria-label="批注类型" value={type} disabled={readOnly||busy} onChange={event=>{setType(event.target.value as PdfAnnotation['type']);setGeometry({x:0,y:0});}}><option value="comment">附注</option><option value="highlight">高亮</option><option value="ink">手绘</option>{type==='text'&&<option value="text">文字附注</option>}</select></label><label>批注<textarea aria-label="批注" value={content} onChange={event=>setContent(event.target.value)} disabled={readOnly||busy} maxLength={20000}/></label>
   {(['x','y'] as const).map(axis=><label key={axis}>PDF {axis.toUpperCase()}<input aria-label={'PDF '+axis.toUpperCase()} type="number" step="any" value={Number.isFinite(geometry[axis])?geometry[axis]:''} disabled={readOnly||busy} onChange={event=>{const value=Number(event.target.value),delta=value-geometry[axis];setGeometry({...geometry,[axis]:value,...(geometry.points?{points:geometry.points.map((n,index)=>index%2===(axis==='x'?0:1)?n+delta:n)}:{})});}}/></label>)}
   {type==='highlight'&&(['width','height'] as const).map(axis=><label key={axis}>{axis==='width'?'批注宽度':'批注高度'}<input aria-label={axis==='width'?'批注宽度':'批注高度'} type="number" min="1" step="any" value={geometry[axis]??''} disabled={readOnly||busy} onChange={event=>setGeometry({...geometry,[axis]:Number(event.target.value)})}/></label>)}
   <p>定位：X {geometry.x.toFixed(1)}，Y {geometry.y.toFixed(1)}{geometry.width?`，宽 ${geometry.width.toFixed(1)}，高 ${geometry.height?.toFixed(1)}`:''}（PDF 坐标）。</p>
   <button className="primary" disabled={readOnly||busy||!validPdfAnnotation(annotation,true)||type==='comment'&&!content.trim()}>{busy?'正在提交…':editing?'保存批注':'添加批注'}</button>{hasDraft&&<button type="button" disabled={busy} onClick={clearDraft}>{editing?'取消修改':'清空批注草稿'}</button>}
  </form>
    {error&&<div role="alert">{error}{pending.current&&!readOnly&&<button disabled={busy} onClick={()=>{const operation=pending.current!;void submit({type:operation.type,annotationId:operation.annotationId,base:operation.base,...(operation.annotation?{annotation:operation.annotation}:{})});}}>重试刚才操作</button>}{!readOnly&&<button disabled={busy} onClick={()=>void latest()}>载入最新修订，保留输入</button>}</div>}
  
    <h3>批注列表 · {annotations.length}</h3><ol>{annotations.map(value=><li key={pdfAnnotationId(value)}>第 {String(value.get('page'))} 页：{String(value.get('content'))}<button onClick={()=>{setCommentTarget(pdfAnnotationId(value));setPage(String(value.get('page')));}}>选择批注供评论</button>{annotationThreads.has(pdfAnnotationId(value))&&<button onClick={()=>{setPage(String(value.get('page')));requestOfficeThread(workspace.id,item.id,annotationThreads.get(pdfAnnotationId(value))![0]);}}>查看讨论 {annotationThreads.get(pdfAnnotationId(value))!.length}</button>}<button disabled={readOnly||busy||hasDraft} onClick={()=>select(value)}>修改批注</button><button disabled={readOnly||busy||hasDraft} onClick={()=>void submit({type:'delete',annotationId:pdfAnnotationId(value),base:pdfAnnotationRevision(value)})}>删除</button></li>)}</ol>
    <OfficePresenceList peers={peers} describe={peer=>peer.focus?.kind==='pdf'?`第 ${peer.focus.page} 页`:undefined} onLocate={locatePeer} following={follow.following} onFollow={follow.follow}/>
    <PdfNativeAnnotations item={item} workspace={workspace} readOnly={readOnly}/>
    <PdfFormPanel item={item} workspace={workspace} onSpecs={setFormSpecs} visibleIds={visibleFields}/>
   </aside>
  </div>
  <OfficeStatusBar><label>页码 <input className="pdf-page-input" aria-label="页码" inputMode="numeric" value={page} disabled={!pages} onChange={event=>setPage(event.target.value)} onBlur={()=>goto(current)} onKeyDown={event=>{if(event.key==='Enter')goto(current);}}/> / {pages||'—'}</label><span>{Math.round(scale*100)}%</span>{rotation?<span>旋转 {rotation}°</span>:null}<span>批注 {annotations.length} 条</span>{peers.length>0&&<span>在线 {peers.length} 人</span>}</OfficeStatusBar>
 </section>;
}
