import {captureWordComment,locateWordComment} from './office-comment-location';
import {useWordFlowSelection} from './use-word-flow-selection';
import {useWordPresence} from './use-word-presence';
import {WordCollaborationLayer} from './word-collaboration-layer';
import type {OfficePeer} from '../../../common/nbstore/src/office-presence';
import {OfficeActivityButton,useFollowPeer} from './office-collaboration-ui';
import {useWordPagination} from './use-word-pagination';
import {OfficePresenceList} from './office-presence-list';
import {wordDomPoint} from './word-flow-dom';
import {WordPrintPreview} from './word-print-preview';
import {WordHeaderDialog} from './word-header-dialog';
import {useWordAuthor,setWordTracking,revisionAt,resolveRevision,WordReviewPane} from './word-review';
import {wordTracking,wordRevisions,setWordTrackingAuthor,deleteWordImageTracked,type WordRevision} from '../../../common/nbstore/src/office-word-revision';
import {wordHeaderLines,wordHeaderScope,wordHeaderImageKey,emptyWordLine,WORD_HEADER_SLOTS,type WordHeaderLine} from '../../../common/nbstore/src/office-word-header';
import {OfficeComments} from './office-comments-ui';
import {wordWritingEnd,wordTextIdentity} from './office-word-flow';
import {restoreWordCaret} from './word-rich-input';
import {WordEditingContext} from './word-editing-context';
import {InsertOfficeImage,WordImageEditor,WordHeaderPicture} from './office-image-ui';
import {WordTableDialog} from './word-table-dialog';
import {wordBlockTexts,type WordBlock} from '../../../common/nbstore/src/office-word-table';
import {WordTableEditor} from './word-table-editor';
import {getWordPage,getWordHeaderFooter,getWordHeaderImages} from './office-word-page';
import {WordPageDialog} from './word-page-dialog';
import {useEffect,useLayoutEffect,useId,useRef,useState} from 'react';
import {Y,LOCAL_ORIGIN} from '@mydoc/crdt';
import type {LocalWorkspace,OfficeWordDocument} from './workspace';
import {useOfficeDocument} from './use-office-document';
import {wordListMarkers} from './office-word-format';
import {WordParagraphEditor} from './word-paragraph-editor';
import {Modal} from './dialogs';
import {cloudRequest,cloudJson,CLOUD_ORIGIN} from '../../../common/nbstore/src/cloud';
import type {WordTableOperation} from '../../../common/nbstore/src/office-word-table-edit';
import {downloadOfficeBytes} from './office-download';
import {SaveStatus} from './text-editor-status';
import {wordDomSelection} from './word-rich-input';
import {readWordFlowDom} from './word-flow-dom';
import {resolveWordFlowSelection} from './word-flow-selection';
import {splitWordParagraph} from './office-word-flow';
import {applyWordFormat,captureWordSelection,wordFormatAt} from './office-word-format';
import {findWordMatches,replaceWordMatch,replaceAllWord,wordOutline,wordCounts} from './word-find';
import {RibbonGroup,RibbonButton} from './office-ribbon';
import {Table,SeparatorHorizontal,CalendarDays,Omega,MessageSquare,ListTree,Printer,Search,Replace,FileText,X,PanelTop,PanelBottom,Hash,PenLine,ChevronUp,ChevronDown,PanelRight,Check,CheckCheck,Ban} from 'lucide-react';
const WORD_SYMBOLS=['，','。','、','；','：','“','”','《','》','【','】','—','…','·','©','®','™','°','±','×','÷','≈','≠','≤','≥','√','∞','→','←','↑','↓','★','☆','●','○','■','□','①','②','③','④','⑤','€','£','¥','§','¶'];



const WORD_TABS=['开始','插入','布局','审阅','视图'];
export function WordEditor({item,workspace,name}:{item:OfficeWordDocument;workspace:LocalWorkspace;name:string}){
  useOfficeDocument(item.doc,workspace);
  const [printOpen,setPrintOpen]=useState(false);
  const [pageOpen,setPageOpen]=useState(false),[tableOpen,setTableOpen]=useState(false),[headerOpen,setHeaderOpen]=useState<'header'|'footer'>();
  const [undo,setUndo]=useState<Y.UndoManager>();
  useEffect(()=>{const manager=new Y.UndoManager(item.paragraphs,{trackedOrigins:new Set([LOCAL_ORIGIN])});setUndo(manager);return()=>manager.destroy();},[item.paragraphs]);
  const ribbonId=useId();
  const [availableWidth,setAvailableWidth]=useState(900);
  const stage=useRef<HTMLDivElement>(null);
  useEffect(()=>{const element=stage.current;if(!element)return;const observer=new ResizeObserver(()=>setAvailableWidth(element.clientWidth));observer.observe(element);return()=>observer.disconnect();},[]);
  const page=getWordPage(item.doc),headerFooter=getWordHeaderFooter(item.doc),headerImages=getWordHeaderImages(item.doc,headerFooter);
  const [tab,setTab]=useState('开始'),[toolbar,setToolbar]=useState<HTMLDivElement|null>(null),[active,setActive]=useState<Y.Text|null>(()=>wordBlockTexts(item.paragraphs.toArray())[0]??null),[zoom,setZoom]=useState<number|'fit'>('fit');
  useEffect(()=>{if(active&&!wordBlockTexts(item.paragraphs.toArray()).includes(active))setActive(wordBlockTexts(item.paragraphs.toArray())[0]??null);});
  const scale=zoom==='fit'?Math.min(1,Math.max(.1,(availableWidth-(availableWidth<700?24:64))/(page.width/15))):zoom/100;
  const [exportOpen,setExportOpen]=useState(false),[exportBusy,setExportBusy]=useState(false),[exportError,setExportError]=useState('');
  const shell=useRef<HTMLElement>(null);
  const keys=useRef(new WeakMap<object,string>());
  // 最终状态 (no markup) is a read-only view: deleted text is hidden, so offsets would not match what is shown.
  const [markup,setMarkup]=useState<'all'|'final'>('all'),[reviewOpen,setReviewOpen]=useState(false),author=useWordAuthor(workspace);
  const documentReadOnly=workspace.documentReadOnly(item.id),readOnly=documentReadOnly||markup==='final',tracking=wordTracking(item.doc)&&author?author:undefined;
  // Formatting helpers look up the tracking author per document, so every formatting path is tracked alike.
  useEffect(()=>{setWordTrackingAuthor(item.doc,tracking);return()=>setWordTrackingAuthor(item.doc,undefined);},[item.doc,tracking?.id,tracking?.name]);
  const texts=wordBlockTexts(item.paragraphs.toArray()),listMarkers=wordListMarkers(texts),markers=new Map(texts.map((text,i)=>[text,listMarkers[i]]));
  let paragraphNumber=0,tableNumber=0;
  const keyFor=(paragraph:WordBlock)=>{let key=keys.current.get(paragraph);if(!key){key=crypto.randomUUID();keys.current.set(paragraph,key);}return key;};
  const pendingFocus=useRef<{text:Y.Text;offset:number}>();
  const placeFocus=()=>{const pending=pendingFocus.current;if(!pending)return;const element=shell.current?.querySelector<HTMLElement>(`.word-rich-input[data-word-text="${wordTextIdentity(pending.text)}"]`);if(!element)return;pendingFocus.current=undefined;element.focus();const offset=Math.min(pending.offset,pending.text.length);restoreWordCaret(element,offset,offset);};
  const focusText=(text:Y.Text,offset=0)=>{pendingFocus.current={text,offset};setActive(text);placeFocus();};
  useLayoutEffect(placeFocus);
  useEffect(()=>{
    if(readOnly)return;
    const blocks=item.paragraphs.toArray();if(blocks.length>1||blocks.length===1&&(!(blocks[0] instanceof Y.Text)||blocks[0].length))return;
    const frame=requestAnimationFrame(()=>{if(document.querySelector('dialog[open]')||shell.current?.closest('.workbench-pane:not(.active-pane)'))return;workspace.assertWritable(item.id);focusText(wordWritingEnd(item.doc));});
    return()=>cancelAnimationFrame(frame);
  },[item.doc,readOnly]);
  const enterPaper=(event:React.MouseEvent<HTMLElement>)=>{
    if(readOnly||event.button!==0||(event.target as HTMLElement).closest('.word-rich-input,.office-word-table,.office-word-image,button,input,select,a'))return;
    event.preventDefault();workspace.assertWritable(item.id);
    const blocks=item.paragraphs.toArray(),paragraphs=blocks.filter((block):block is Y.Text=>block instanceof Y.Text),inputs=[...event.currentTarget.querySelectorAll<HTMLElement>(':scope > .office-word-paragraph .word-rich-input')];
    for(let i=0;i<inputs.length;i++){const rect=inputs[i].getBoundingClientRect();if(event.clientY<=rect.bottom){focusText(paragraphs[i],event.clientY<rect.top?0:paragraphs[i].length);return;}}
    const text=wordWritingEnd(item.doc);focusText(text,text.length);
  };
  // Team workspaces: table rows, columns and merges are changed by the server so concurrent edits stay rectangular.
  const tableRemote=async(operation:WordTableOperation)=>{workspace.assertWritable(item.id);await workspace.flushResource(item.id);const result=await cloudJson<{update:string;focus?:string}>('/office-word-table-operation?'+new URLSearchParams({workspace:workspace.id,id:item.id}),'POST',operation).catch(async error=>{if((error as {status?:number}).status===409){try{const response=await cloudRequest('/document?'+new URLSearchParams({workspace:workspace.id,id:item.id}));Y.applyUpdate(item.doc,new Uint8Array(await response.arrayBuffer()),CLOUD_ORIGIN);}catch{}}throw error;});Y.applyUpdate(item.doc,Uint8Array.from(atob(result.update),c=>c.charCodeAt(0)),CLOUD_ORIGIN);return result.focus;};
  const flow=useWordFlowSelection(shell,item.doc,readOnly,()=>workspace.assertWritable(item.id),focusText,()=>undo?.stopCapturing(),tracking);
  const presence=useWordPresence(shell,item.doc,!!workspace.sync),paperRef=useRef<HTMLElement>(null),pageBreaks=useWordPagination(paperRef,page);
  // Like Word, a header or footer taller than its page margin pushes the body text away instead of covering it.
  const [marginGrow,setMarginGrow]=useState({top:0,bottom:0}),marginKey=JSON.stringify([headerFooter??null,headerImages,page]);
  useLayoutEffect(()=>{const paper=paperRef.current;if(!paper)return;
   const strips=[...paper.querySelectorAll<HTMLElement>('.word-margin-layer>.word-margin-text')];
   const extent=(part:'header'|'footer',margin:number)=>{const strip=strips.find(item=>item.parentElement?.dataset.part===part);if(!strip||strip.dataset.empty)return 0;const offset=parseFloat(part==='header'?strip.style.top:strip.style.bottom)||0;return Math.max(0,Math.ceil(offset+strip.offsetHeight+4-margin/15));};
   const measure=()=>setMarginGrow(current=>{const next={top:extent('header',page.top),bottom:extent('footer',page.bottom)};return next.top===current.top&&next.bottom===current.bottom?current:next;});
   measure();const observer=new ResizeObserver(measure);for(const strip of strips)observer.observe(strip);return()=>observer.disconnect();
  },[marginKey]);
  const [painter,setPainter]=useState<Record<string,unknown>>(),[find,setFind]=useState<{query:string;replacement:string;matchCase:boolean;index:number;replace:boolean}>(),[outline,setOutline]=useState(false),[countOpen,setCountOpen]=useState(false),[notice,setNotice]=useState('');
  const findInput=useRef<HTMLInputElement>(null),caret=useRef<{text:Y.Text;offset:number}>();
  // Remember the last caret in the document so ribbon controls that take focus can still insert there.
  useEffect(()=>{const track=()=>{const dom=window.getSelection(),element=dom?.anchorNode?(dom.anchorNode instanceof HTMLElement?dom.anchorNode:dom.anchorNode.parentElement)?.closest<HTMLElement>('.word-rich-input'):null;if(!element||!shell.current?.contains(element))return;const text=wordBlockTexts(item.paragraphs.toArray()).find(value=>wordTextIdentity(value)===element.dataset.wordText),range=wordDomSelection(element);if(text&&range)caret.current={text,offset:range.end};};document.addEventListener('selectionchange',track);return()=>document.removeEventListener('selectionchange',track);},[item.paragraphs]);
  const writable=()=>{if(readOnly)throw new Error('当前文档只读');workspace.assertWritable(item.id);};
  const insertText=(value:string)=>{try{writable();const at=caret.current;if(!at||!at.text.doc)throw new Error('请先在正文中放置光标');const offset=Math.min(at.offset,at.text.length),format=offset>0?wordFormatAt(at.text,offset-1):{};undo?.stopCapturing();item.doc.transact(()=>at.text.insert(offset,value,format),LOCAL_ORIGIN);caret.current={text:at.text,offset:offset+value.length};focusText(at.text,offset+value.length);setNotice('');}catch(error){setNotice(error instanceof Error?error.message:'插入失败');}};
  const insertPageBreak=()=>{try{writable();const at=caret.current;if(!at||!at.text.doc)throw new Error('请先在正文中放置光标');if(!item.paragraphs.toArray().includes(at.text))throw new Error('表格单元格内不能插入分页符');undo?.stopCapturing();const next=splitWordParagraph(at.text,Math.min(at.offset,at.text.length));item.doc.transact(()=>next.setAttribute('pageBreakBefore',true),LOCAL_ORIGIN);caret.current={text:next,offset:0};focusText(next,0);setNotice('');}catch(error){setNotice(error instanceof Error?error.message:'插入分页符失败');}};
  // Format painter: the next non-empty selection receives the copied character format, then the painter turns off.
  useEffect(()=>{if(!painter)return;const apply=()=>{const root=shell.current,dom=window.getSelection();if(!root||!dom?.rangeCount||dom.isCollapsed)return;try{writable();const flow=readWordFlowDom(root,item.doc);
    if(flow){const {ranges}=resolveWordFlowSelection(item.doc,flow);item.doc.transact(()=>{for(const range of ranges)if(range.end>range.start)applyWordFormat(range.text,captureWordSelection(range.text,range.start,range.end)!,painter);},LOCAL_ORIGIN);}
    else{const element=(dom.anchorNode instanceof HTMLElement?dom.anchorNode:dom.anchorNode?.parentElement)?.closest<HTMLElement>('.word-rich-input'),text=element?wordBlockTexts(item.paragraphs.toArray()).find(value=>wordTextIdentity(value)===element.dataset.wordText):undefined,range=element&&wordDomSelection(element);if(!text||!range||range.end<=range.start)return;applyWordFormat(text,captureWordSelection(text,range.start,range.end)!,painter);}
    setPainter(undefined);setNotice('');}catch(error){setNotice(error instanceof Error?error.message:'格式刷应用失败');setPainter(undefined);}};
   document.addEventListener('mouseup',apply);document.addEventListener('keyup',apply);return()=>{document.removeEventListener('mouseup',apply);document.removeEventListener('keyup',apply);};},[painter]);
  const matches=find?findWordMatches(item.doc,find.query,find.matchCase):[];
  const showMatch=(index:number)=>{const match=matches[index];if(!match)return;const element=shell.current?.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(match.text)}"]`);if(!element)return;element.focus();const range=document.createRange();range.setStart(...wordDomPoint(element,match.start));range.setEnd(...wordDomPoint(element,match.end));const dom=window.getSelection();dom?.removeAllRanges();dom?.addRange(range);element.scrollIntoView({block:'center'});};
  const findNext=()=>{if(!find?.query)return;if(!matches.length){setNotice('未找到“'+find.query+'”');return;}setNotice('');const index=(find.index+1)%matches.length;setFind({...find,index});showMatch(index);};
  const replaceOne=()=>{if(!find?.query)return;try{writable();const current=matches[find.index];if(!current){findNext();return;}undo?.stopCapturing();replaceWordMatch(item.doc,current,find.replacement,tracking);setFind({...find,index:find.index-1});setNotice('已替换 1 处');}catch(error){setNotice(error instanceof Error?error.message:'替换失败');}};
  const replaceAll=()=>{if(!find?.query)return;try{writable();undo?.stopCapturing();const count=replaceAllWord(item.doc,find.query,find.replacement,find.matchCase,tracking);undo?.stopCapturing();setFind({...find,index:-1});setNotice(count?`已全部替换，共 ${count} 处`:'未找到可替换的内容');}catch(error){setNotice(error instanceof Error?error.message:'替换失败');}};
  const openFind=(replace:boolean)=>{setFind(value=>({query:value?.query??'',replacement:value?.replacement??'',matchCase:value?.matchCase??false,index:-1,replace}));requestAnimationFrame(()=>findInput.current?.focus());};
  const headings=outline?wordOutline(item.doc):[],counts=wordCounts(item.doc),revisionCount=new Set(wordRevisions(item.doc).map(revision=>revision.id)).size;
  /** Select a revision's text (scrolling to it) so 接受/拒绝 and the pane act on what the user sees. */
  const showRevision=(revision:WordRevision)=>{
   // Pictures have no text: select and show the figure.
   if(!revision.text){const id=revision.block?._item?.id,figure=id&&shell.current?.querySelector<HTMLElement>(`[data-word-image="${id.client}:${id.clock}"]`);if(figure){figure.focus();figure.scrollIntoView({block:'center'});}return;}
   const text=revision.text,element=shell.current?.querySelector<HTMLElement>(`[data-word-text="${wordTextIdentity(text)}"]`);if(!element)return;element.focus();const range=document.createRange();range.setStart(...wordDomPoint(element,revision.start));range.setEnd(...wordDomPoint(element,revision.end));const dom=window.getSelection();dom?.removeAllRanges();dom?.addRange(range);caret.current={text,offset:revision.start};element.scrollIntoView({block:'center'});};
  const step=(direction:1|-1)=>{const next=revisionAt(item.doc,caret.current,direction,true);if(next)showRevision(next);};
  const decide=(accept:boolean)=>{try{writable();const current=revisionAt(item.doc,caret.current);if(!current)return;undo?.stopCapturing();resolveRevision(item.doc,accept,current.id);undo?.stopCapturing();const next=revisionAt(item.doc,caret.current);if(next)showRevision(next);setNotice(accept?'已接受修订':'已拒绝修订');}catch(error){setNotice(error instanceof Error?error.message:'修订处理失败');}};
  const locatePeer=(peer:OfficePeer)=>{const target=presence.resolve(peer),identity=target?.text._item?`${target.text._item.id.client}:${target.text._item.id.clock}`:undefined,node=identity?shell.current?.querySelector<HTMLElement>(`[data-word-text="${identity}"]`):null;if(!target||!node)return;node.focus();const range=document.createRange();range.setStart(...wordDomPoint(node,target.start));range.setEnd(...wordDomPoint(node,target.end));const selection=window.getSelection();selection?.removeAllRanges();selection?.addRange(range);node.scrollIntoView({block:'center'});};
  const follow=useFollowPeer(presence.peers,locatePeer,shell);
  // The paragraph or cell holding a remote caret is also marked; exact carets and selections are drawn by WordCollaborationLayer.
  useEffect(()=>{const element=shell.current;if(!element)return;for(const node of element.querySelectorAll<HTMLElement>('.word-rich-input[data-peer]')){node.removeAttribute('data-peer');node.style.removeProperty('--peer-color');node.removeAttribute('title');}for(const peer of presence.peers){const target=presence.resolve(peer),identity=target&&target.text._item?`${target.text._item.id.client}:${target.text._item.id.clock}`:undefined,node=identity?element.querySelector<HTMLElement>(`[data-word-text="${identity}"]`):null;if(node&&!node.dataset.peer){node.dataset.peer=peer.name;node.style.setProperty('--peer-color',peer.color);node.title=peer.name+' 正在此处编辑';}}});
  return <WordEditingContext.Provider value={{active,activate:setActive,toolbar,focus:focusText,flowActive:flow.active,formatFlow:flow.format,painter,setPainter,tracking}}><section ref={shell} className="office-word-editor office-document-shell" aria-label={name} data-painter={painter?'':undefined} data-markup={markup} onKeyDownCapture={event=>{if((event.metaKey||event.ctrlKey)&&!event.shiftKey&&['f','h'].includes(event.key.toLowerCase())){event.preventDefault();event.stopPropagation();openFind(event.key.toLowerCase()==='h');return;}if(event.key==='Escape'&&painter){setPainter(undefined);return;}if(event.target instanceof HTMLElement&&event.target.isContentEditable&&(event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==='z'){event.preventDefault();event.stopPropagation();if(!readOnly){workspace.assertWritable(item.id);if(event.shiftKey)undo?.redo();else undo?.undo();}}}}>
    <header className="office-document-title"><div className="office-file-badge">W</div><div><h2>{name}</h2><span>{documentReadOnly?'只读':markup==='final'?'无标记视图（只读）':workspace.sync?'协同编辑':'本地编辑'}{wordTracking(item.doc)?' · 修订已打开':''}</span></div><SaveStatus persistence={item.persistence} sync={workspace.sync}/><div className="word-undo-tools"><button disabled={readOnly||!undo?.canUndo()} onMouseDown={event=>event.preventDefault()} onClick={()=>{workspace.assertWritable(item.id);undo?.undo();}}>撤销</button><button disabled={readOnly||!undo?.canRedo()} onMouseDown={event=>event.preventDefault()} onClick={()=>{workspace.assertWritable(item.id);undo?.redo();}}>重做</button></div><button onClick={()=>setPrintOpen(true)}>打印预览</button><button className="office-export" onClick={()=>{setExportError('');setExportOpen(true);}}>导出 DOCX 正文副本</button><OfficeActivityButton doc={item.doc} workspace={workspace} id={item.id}/><OfficeComments doc={item.doc} workspace={workspace} id={item.id} captureAnchor={()=>captureWordComment(shell.current,item.doc)} onLocate={anchor=>locateWordComment(shell.current,item.doc,anchor)}/></header>
    <OfficePresenceList peers={presence.peers} describe={peer=>{const target=presence.resolve(peer);return target?`第 ${target.index+1} 处正文`+(target.end>target.start?`，选中 ${target.end-target.start} 字`:''):undefined;}} onLocate={locatePeer} following={follow.following} onFollow={follow.follow}/>
    <div className="office-ribbon-tabs" role="tablist" aria-label="Word 功能区">{WORD_TABS.map((name,i)=><button key={name} role="tab" aria-selected={tab===name} aria-controls={ribbonId} tabIndex={tab===name?0:-1} onKeyDown={event=>{if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return;event.preventDefault();const next=event.key==='Home'?0:event.key==='End'?WORD_TABS.length-1:(i+(event.key==='ArrowRight'?1:WORD_TABS.length-1))%WORD_TABS.length;setTab(WORD_TABS[next]);(event.currentTarget.parentElement?.children[next] as HTMLElement)?.focus();}} onClick={()=>setTab(name)}>{name}</button>)}</div>
    <div id={ribbonId} className="office-ribbon-panel" role="tabpanel" aria-label={tab} data-tab={tab}>
<div ref={setToolbar} className="word-format-slot" hidden={tab!=='开始'&&tab!=='布局'}/>
      {tab==='开始'&&!active&&<p className="organization-help">点击纸张即可开始输入。</p>}
      {tab==='开始'&&<div className="office-ribbon-groups word-edit-group"><RibbonGroup label="编辑"><RibbonButton icon={Search} label="查找" onClick={()=>openFind(false)}/><RibbonButton icon={Replace} label="替换" disabled={readOnly} onClick={()=>openFind(true)}/></RibbonGroup></div>}
      {tab==='插入'&&<div className="office-ribbon-groups"><RibbonGroup label="表格与图片"><RibbonButton icon={Table} label="新建表格" large disabled={readOnly} onClick={()=>setTableOpen(true)}/><InsertOfficeImage workspace={workspace} id={item.id} doc={item.doc} target={item.paragraphs} kind="word" disabled={readOnly}/></RibbonGroup>
       <RibbonGroup label="页"><RibbonButton icon={SeparatorHorizontal} label="分页符" disabled={readOnly} onMouseDown={event=>event.preventDefault()} onClick={insertPageBreak}/></RibbonGroup>
       <RibbonGroup label="页眉和页脚"><RibbonButton icon={PanelTop} label="页眉" onClick={()=>setHeaderOpen('header')}/><RibbonButton icon={PanelBottom} label="页脚" onClick={()=>setHeaderOpen('footer')}/><RibbonButton icon={Hash} label="页码" onClick={()=>setHeaderOpen('footer')}/></RibbonGroup>
       <RibbonGroup label="文本"><RibbonButton icon={CalendarDays} label="日期和时间" disabled={readOnly} onMouseDown={event=>event.preventDefault()} onClick={()=>insertText(new Date().toLocaleDateString('zh-CN',{year:'numeric',month:'long',day:'numeric'}))}/><label className="ribbon-color" title="符号"><Omega size={15} aria-hidden="true"/><select aria-label="插入符号" value="" disabled={readOnly} onChange={event=>{if(event.target.value)insertText(event.target.value);}}><option value="">符号</option>{WORD_SYMBOLS.map(symbol=><option key={symbol} value={symbol}>{symbol}</option>)}</select></label></RibbonGroup></div>}
      {tab==='布局'&&<div className="office-insert-tools"><button onClick={()=>setPageOpen(true)}>页面设置</button><span aria-label="当前页面设置">{page.width>page.height?'横向':'纵向'} · {Math.round(page.width/20)} × {Math.round(page.height/20)} 磅</span></div>}
      {tab==='审阅'&&<div className="office-ribbon-groups"><RibbonGroup label="批注"><RibbonButton icon={MessageSquare} label="新建批注" onMouseDown={event=>event.preventDefault()} onClick={()=>[...shell.current?.querySelectorAll<HTMLButtonElement>('.office-document-title button')??[]].find(button=>button.textContent?.startsWith('评论'))?.click()}/></RibbonGroup><RibbonGroup label="校对"><RibbonButton icon={FileText} label="字数统计" onClick={()=>setCountOpen(true)}/></RibbonGroup>
       <RibbonGroup label="修订"><RibbonButton icon={PenLine} label="修订" large pressed={wordTracking(item.doc)} disabled={documentReadOnly||!author} title={wordTracking(item.doc)?'正在记录修订：插入与删除会保留标记，等待接受或拒绝':'开始记录修订'} onClick={()=>{workspace.assertWritable(item.id);setWordTracking(item.doc,!wordTracking(item.doc));}}/>
        <select aria-label="显示标记" value={markup} onChange={event=>setMarkup(event.target.value as 'all'|'final')}><option value="all">所有标记</option><option value="final">无标记（只读）</option></select><RibbonButton icon={PanelRight} label="审阅窗格" pressed={reviewOpen} onClick={()=>setReviewOpen(value=>!value)}/></RibbonGroup>
       <RibbonGroup label="更改"><RibbonButton icon={Check} label="接受" disabled={readOnly||!revisionCount} onMouseDown={event=>event.preventDefault()} onClick={()=>decide(true)}/><RibbonButton icon={Ban} label="拒绝" disabled={readOnly||!revisionCount} onMouseDown={event=>event.preventDefault()} onClick={()=>decide(false)}/>
        <RibbonButton icon={ChevronUp} label="上一处修订" disabled={!revisionCount} onMouseDown={event=>event.preventDefault()} onClick={()=>step(-1)}/><RibbonButton icon={ChevronDown} label="下一处修订" disabled={!revisionCount} onMouseDown={event=>event.preventDefault()} onClick={()=>step(1)}/>
        <RibbonButton icon={CheckCheck} label="接受所有修订" disabled={readOnly||!revisionCount} onClick={()=>{workspace.assertWritable(item.id);undo?.stopCapturing();const n=resolveRevision(item.doc,true);setNotice(`已接受 ${n} 处修订`);}}/><RibbonButton icon={X} label="拒绝所有修订" disabled={readOnly||!revisionCount} onClick={()=>{workspace.assertWritable(item.id);undo?.stopCapturing();const n=resolveRevision(item.doc,false);setNotice(`已拒绝 ${n} 处修订`);}}/></RibbonGroup></div>}
      {tab==='视图'&&<div className="office-ribbon-groups"><RibbonGroup label="显示"><RibbonButton icon={ListTree} label="导航窗格" pressed={outline} onClick={()=>setOutline(value=>!value)}/><RibbonButton icon={Printer} label="打印预览" onClick={()=>setPrintOpen(true)}/></RibbonGroup><RibbonGroup label="缩放"><select aria-label="视图缩放" value={zoom} onChange={event=>setZoom(event.target.value==='fit'?'fit':Number(event.target.value))}><option value="fit">适应宽度</option>{[50,75,100,125,150].map(value=><option key={value} value={value}>{value}%</option>)}</select><RibbonButton label="100%" pressed={zoom===100} onClick={()=>setZoom(100)}/></RibbonGroup></div>}
    </div>
    {find&&<div className="word-find" role="search" aria-label="查找和替换"><label>查找<input ref={findInput} aria-label="查找内容" value={find.query} onChange={event=>setFind({...find,query:event.target.value,index:-1})} onKeyDown={event=>{if(event.key==='Enter'){event.preventDefault();findNext();}if(event.key==='Escape')setFind(undefined);}}/></label>{find.replace&&<label>替换为<input aria-label="替换为" value={find.replacement} onChange={event=>setFind({...find,replacement:event.target.value})}/></label>}<label><input type="checkbox" checked={find.matchCase} onChange={event=>setFind({...find,matchCase:event.target.checked,index:-1})}/>区分大小写</label><span role="status">{find.query?(matches.length>5000?'超过 5000 处':`${find.index>=0?find.index+1+' / ':''}${matches.length} 处`):''}</span><button onClick={findNext}>查找下一处</button>{find.replace?<><button disabled={readOnly||!find.query} onClick={replaceOne}>替换</button><button disabled={readOnly||!find.query} onClick={replaceAll}>全部替换</button></>:<button disabled={readOnly} onClick={()=>setFind({...find,replace:true})}>替换…</button>}<button aria-label="关闭查找" onClick={()=>setFind(undefined)}><X size={14} aria-hidden="true"/></button></div>}
    {(notice||painter)&&<p className="word-notice" role="status">{painter?'格式刷：选中要应用格式的文字，按 Esc 取消':notice}</p>}
    {countOpen&&<Modal title="字数统计" onClose={()=>setCountOpen(false)}><dl className="word-count-list"><dt>页数（估算）</dt><dd>{pageBreaks.length+1}</dd><dt>字数</dt><dd>{counts.words}</dd><dt>字符数（不计空格）</dt><dd>{counts.charactersNoSpaces}</dd><dt>字符数（计空格）</dt><dd>{counts.characters}</dd><dt>段落数</dt><dd>{counts.paragraphs}</dd></dl><p>中日韩文字每个字计一字，其他连续字母或数字计一个词；包含表格内文字。</p><div className="dialog-actions"><button onClick={()=>setCountOpen(false)}>关闭</button></div></Modal>}
    {tableOpen&&<WordTableDialog doc={item.doc} readOnly={readOnly} assertWritable={()=>workspace.assertWritable(item.id)} onClose={()=>setTableOpen(false)}/>}
    {printOpen&&shell.current?.querySelector<HTMLElement>('.word-paper')&&<WordPrintPreview source={shell.current.querySelector<HTMLElement>('.word-paper')!} page={page} headerFooter={headerFooter} headerImages={headerImages} workspace={workspace} id={item.id} name={name} onClose={()=>setPrintOpen(false)}/>}
    {pageOpen&&<WordPageDialog doc={item.doc} readOnly={readOnly} onClose={()=>setPageOpen(false)}/>}
    {headerOpen&&<WordHeaderDialog doc={item.doc} workspace={workspace} id={item.id} readOnly={readOnly} focus={headerOpen} onClose={()=>setHeaderOpen(undefined)}/>}
    {exportOpen&&<Modal title="导出 DOCX 正文兼容副本" onClose={()=>{if(!exportBusy)setExportOpen(false);}}>
      <p>导出当前已保存的纸张尺寸、方向、页边距、段前分页、基础矩形表格、段落文字、一级至三级标题、单层项目符号和编号列表、网页超链接、段落对齐、左右/首行/悬挂缩进、行距与段前段后间距、加粗、斜体、单下划线、删除线、上下标、突出显示、字体、字号、RGB/自动文字颜色、合并单元格、空白、制表符、换行及独立 PNG/JPEG 图片。文件评论写为 Word 批注（回复成线程、已解决的标为完成，定位失效的放在文首）。原件中的其他样式、复杂表格、页眉页脚和修订不会包含在副本中，原件保持不变。</p>
      {exportError&&<p role="alert" className="form-error">{exportError}</p>}
      <div className="dialog-actions"><button disabled={exportBusy} onClick={()=>setExportOpen(false)}>取消</button><button className="primary" disabled={exportBusy} onClick={()=>{
        setExportBusy(true);setExportError('');
        void (async()=>{
          await workspace.flushResource(item.id);
          const response=await cloudRequest('/office-export?'+new URLSearchParams({workspace:workspace.id,id:item.id,format:'docx',scope:'paragraphs'}));
          if(!response.ok){const result=await response.json().catch(()=>null);throw new Error(result?.error??'导出失败，请重试');}
          downloadOfficeBytes(name.replace(/\.[^.]+$/,'')+'-正文副本.docx',new Uint8Array(await response.arrayBuffer()),'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          setExportOpen(false);
        })().catch(error=>setExportError(error instanceof Error?error.message:'导出失败，请重试')).finally(()=>setExportBusy(false));
      }}>{exportBusy?'正在生成…':'确认导出正文副本'}</button></div>
    </Modal>}
    {flow.error&&<p role="alert">{flow.error}</p>}<div className="word-body" data-outline={outline||undefined}>{outline&&<nav className="word-outline" aria-label="导航窗格"><h3>导航</h3>{headings.length?<ol>{headings.map((entry,i)=><li key={i} data-level={entry.level}><button title={entry.title} aria-current={active===entry.text?'location':undefined} onClick={()=>{focusText(entry.text,0);shell.current?.querySelector(`[data-word-text="${wordTextIdentity(entry.text)}"]`)?.scrollIntoView({block:'start'});}}>{entry.title}</button></li>)}</ol>:<p>文档中没有标题。在“开始”中把段落设为标题 1–3 后会显示在这里。</p>}</nav>}<div ref={stage} className="word-document-stage"><div className="word-page-ruler" aria-hidden="true" style={{width:page.width/15*scale}}>{Array.from({length:Math.floor(page.width/1440*2.54)+1},(_,cm)=><span key={cm} style={{left:cm*96/2.54*scale}}>{cm}</span>)}</div><article ref={paperRef} className="word-paper" aria-label="文档正文" onMouseDown={enterPaper} style={{width:page.width/15,minHeight:page.height/15,padding:`${page.top/15+marginGrow.top}px ${page.right/15}px ${page.bottom/15+marginGrow.bottom}px ${page.left/15}px`,zoom:scale}}><WordCollaborationLayer paper={paperRef} doc={item.doc} workspace={workspace.id} id={item.id} peers={presence.peers} resolve={presence.resolve}/>{(()=>{
      // Header on the first page's top margin and footer on the last page's bottom margin; print and DOCX repeat them per page.
      const pages=pageBreaks.length+1,top=wordHeaderLines(headerFooter,1,pages),bottom=wordHeaderLines(headerFooter,pages,pages),gap=(margin:number)=>Math.max(4,Math.min(720,margin-360)/15);
      // A slot shows its picture above its text; the variant (first/even/default) follows the page number.
      const strip=(line:WordHeaderLine|undefined,part:'header'|'footer',margin:number,pageNumber:number)=>{
       const scope=wordHeaderScope(headerFooter,pageNumber),pictures=WORD_HEADER_SLOTS.map(slot=>headerImages[wordHeaderImageKey(scope,part,slot)]),filled=!!line&&(!emptyWordLine(line)||pictures.some(Boolean));
       return <div className="word-margin-layer" data-part={part}><button type="button" className="word-margin-text" style={{[part==='header'?'top':'bottom']:gap(margin),left:page.left/15,right:page.right/15}} aria-label={part==='header'?'编辑页眉':'编辑页脚'} title={part==='header'?'编辑页眉':'编辑页脚'} onMouseDown={event=>event.stopPropagation()} onClick={()=>setHeaderOpen(part)} data-empty={!filled||undefined}>
        {filled?WORD_HEADER_SLOTS.map((slot,i)=><span key={slot}>{pictures[i]&&<WordHeaderPicture workspace={workspace} id={item.id} picture={pictures[i]!}/>}{line![slot]}</span>):<span className="word-margin-placeholder">{part==='header'?'页眉':'页脚'}</span>}</button></div>;
      };
      return <>{strip(top?.header,'header',page.top,1)}{strip(bottom?.footer,'footer',page.bottom,pages)}</>;
     })()}{item.paragraphs.toArray().map((block,index)=>{
      const remove=()=>{workspace.assertWritable(item.id);if(active===block)setActive(texts.find(text=>text!==block)??null);const current=item.paragraphs.toArray().indexOf(block);if(current<0)return;item.doc.transact(()=>{
       // While tracking, a picture is marked deleted (or, if it is the author's own pending insertion, removed).
       if(tracking&&block instanceof Y.Map&&!deleteWordImageTracked(block,tracking))return;item.paragraphs.delete(current,1);},LOCAL_ORIGIN);};
      return block instanceof Y.Text?<WordParagraphEditor assertWritable={()=>workspace.assertWritable(item.id)} marker={markers.get(block)??''} key={keyFor(block)} paragraph={block} index={paragraphNumber++} readOnly={readOnly}/>:block.get('kind')==='image'?<WordImageEditor key={keyFor(block)} block={block} workspace={workspace} id={item.id} disabled={readOnly} onDelete={remove}/>:<WordTableEditor key={keyFor(block)} remote={workspace.sync?tableRemote:undefined} table={block} index={tableNumber++} readOnly={readOnly} assertWritable={()=>workspace.assertWritable(item.id)} onDelete={remove} markers={markers}/>;
    })}{pageBreaks.map(entry=><div key={entry.page} className="word-page-break-marker" aria-hidden="true" style={{top:entry.top}} data-page={entry.page}><span>第 {entry.page} 页</span></div>)}</article></div>{reviewOpen&&<WordReviewPane doc={item.doc} readOnly={readOnly} onLocate={showRevision} onClose={()=>setReviewOpen(false)}/>}</div><footer className="office-document-status">{(wordTracking(item.doc)||revisionCount>0)&&<button type="button" className="word-status-revisions" onClick={()=>setReviewOpen(true)}>{wordTracking(item.doc)?'修订：打开':'修订：关闭'} · {revisionCount} 处</button>}<span aria-label="页数">共 {pageBreaks.length+1} 页（按浏览器排版估算）</span><button className="word-status-count" onClick={()=>setCountOpen(true)}>字数 {counts.words}</button><span>{texts.reduce((count,text)=>count+Array.from(text.toString()).length,0)} 字符</span><span>{readOnly?'只读':'编辑'} · {page.width>page.height?'横向':'纵向'}</span><label>缩放 <select aria-label="页面缩放" value={zoom} onChange={event=>setZoom(event.target.value==='fit'?'fit':Number(event.target.value))}><option value="fit">适应宽度</option>{[50,75,100,125,150].map(value=><option key={value} value={value}>{value}%</option>)}</select></label></footer></section></WordEditingContext.Provider>;
}
