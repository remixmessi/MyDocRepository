import {officeAnchorTarget} from '../../../common/nbstore/src/office-comment-anchor';
import {SpreadsheetDetachedDrafts} from './spreadsheet-detached-drafts';
import {sheetColumnIds,sheetRowIdentity} from '../../../common/nbstore/src/sheet-columns';
import {OfficeComments} from './office-comments-ui';
import {OfficeActivityButton,useFollowPeer,requestOfficeThread} from './office-collaboration-ui';
import {validOfficeComment,type OfficeCommentThread} from '../../../common/nbstore/src/office-comments';
import {WorksheetControls} from './worksheet-controls';
import {workbookSheets,workbookFormulaSnapshot} from '../../../common/nbstore/src/office-workbook';
import {SpreadsheetPasteDialog} from './spreadsheet-paste-dialog';
import {validXlsxSource} from '../../../common/nbstore/src/office-xlsx-source';
import {spreadsheetCellType} from '../../../common/nbstore/src/spreadsheet-cell-type';
import {DelimitedExportDialog} from './delimited-export-dialog';
import {useEffect,useRef,useState} from 'react';
import {Download,X,Check} from 'lucide-react';
import {structureRevision,type SheetFormatPatch,type SheetOperation} from '../../../common/nbstore/src/spreadsheet-operation';
import {useOfficeDocument} from './use-office-document';
import {SpreadsheetCell,type SheetCellHandle,type SheetCellKey} from './spreadsheet-cell';
import type {LocalWorkspace,OfficeSpreadsheetDocument} from './workspace';
import {previewFormulaWorkbook} from '../../../common/nbstore/src/spreadsheet-formula';
import {SpreadsheetPendingOperations} from './spreadsheet-pending-operations';
import {useOfficePresence} from './use-office-presence';
import {OfficePresenceList} from './office-presence-list';
import {resolveSheetMerges,sheetCellFormat,formatSheetValue,SHEET_FORMAT_LIMITS,type SheetCellFormat} from '../../../common/nbstore/src/spreadsheet-format';
import {cellDraftKey} from '../../../common/nbstore/src/spreadsheet-operation';
import {OfficeTitleBar,OfficeMenu,OfficeStatusBar} from './office-ribbon';
import {SpreadsheetRibbon,SHEET_FUNCTIONS,type SheetView,type SpreadsheetActions} from './spreadsheet-ribbon';
import {SpreadsheetExportDialog} from './spreadsheet-export-dialog';
import {columnName,cellName,parseCellRange,selectionStats,formatStat,autoSumPlan,sortedRows,findMatches,replacePlan,rangeToTsv,cellEdits,type SheetRange} from './spreadsheet-commands';

const rowsReady=(item:{sheet:{length:number}})=>item.sheet.length>0;
export function SpreadsheetDocumentEditor({item:resource,workspace,name}:{item:OfficeSpreadsheetDocument;workspace:LocalWorkspace;name:string}){
 useOfficeDocument(resource.doc,workspace);
 const [sheetId,setSheetId]=useState('first'),sheets=workbookSheets(resource.doc),selectedSheet=sheets.find(sheet=>sheet.id===sheetId)??sheets[0],item={...resource,sheet:selectedSheet.sheet,sheetId:selectedSheet.id};
 const [pasteOpen,setPasteOpen]=useState(false),[delimitedOpen,setDelimitedOpen]=useState(false),[exportKind,setExportKind]=useState<'text'|'formula'>();
 const [operationError,setOperationError]=useState(''),[notice,setNotice]=useState(''),[busy,setBusy]=useState(false);
 const [selected,setSelected]=useState<[number,number]>([0,0]),[anchor,setAnchor]=useState<[number,number]>([0,0]),[reported,setHandle]=useState<SheetCellHandle>();
 const [view,setView]=useState<SheetView>({freezeRow:false,freezeColumn:false,gridlines:true,formulas:false,filter:false}),[filters,setFilters]=useState<Record<string,string>>({}),[zoom,setZoom]=useState(100);
 const [nameDraft,setNameDraft]=useState<string>(),[find,setFind]=useState<{query:string;replacement:string;index:number}>(),[painter,setPainter]=useState<SheetCellFormat>();
 // A handle reported before the selection moved belongs to another cell; never let the formula bar or ribbon edit it.
 const activeKey=rowsReady(item)?cellDraftKey(item.sheet,selected[0],selected[1]):'',handle=reported&&reported.id===activeKey?reported:undefined;
 const readOnly=workspace.documentReadOnly(item.id),disabled=readOnly||busy,dragging=useRef(false),keepAnchor=useRef(false),findInput=useRef<HTMLInputElement>(null),selection=useRef({anchor,selected});selection.current={anchor,selected};
 const source=item.doc.share.has('sheetXlsx')?item.doc.getMap('sheetXlsx').toJSON():undefined;
 const rows=item.sheet.toArray(),columns=sheetColumnIds(item.sheet),width=columns.length,grid=rows.map(row=>Array.from({length:width},(_,c)=>row.get(c)??''));
 const merges=resolveSheetMerges(item.doc,selectedSheet.id,item.sheet),hidden=(r:number,c:number)=>{const owner=merges.owner.get(r+':'+c);return !!owner&&owner!==r+':'+c&&!item.drafts.has(cellDraftKey(item.sheet,r,c));};
 const range:SheetRange={top:Math.min(anchor[0],selected[0]),left:Math.min(anchor[1],selected[1]),bottom:Math.max(anchor[0],selected[0]),right:Math.max(anchor[1],selected[1])},inRange=(r:number,c:number)=>r>=range.top&&r<=range.bottom&&c>=range.left&&c<=range.right;
 const select=(r:number,c:number,extend=false)=>{setSelected([r,c]);if(!extend)setAnchor([r,c]);};
 const sheetName=selectedSheet.name,{book}=workbookFormulaSnapshot(item.doc),byName=new Map(sheets.map(sheet=>[sheet.name,sheet.id])),results=previewFormulaWorkbook(book,(name,row,column)=>spreadsheetCellType(item.doc,row,column,byName.get(name)));
 const isFormula=(r:number,c:number)=>spreadsheetCellType(item.doc,r,c,selectedSheet.id)!=='text'&&(grid[r]?.[c]??'').trim().startsWith('=');
 const shown=(r:number,c:number)=>isFormula(r,c)?String(results[sheetName]?.[r]?.[c]??''):grid[r]?.[c]??'';
 const display=(r:number,c:number)=>{const format=sheetCellFormat(item.doc,selectedSheet.id,item.sheet,r,c);return formatSheetValue(shown(r,c),format?.numberFormat);};
 // Filters hide data rows (the first row stays as the header) and are a local view, not shared content.
 const active=Object.entries(filters).filter(([,text])=>text.trim()),visible=(r:number)=>r===0||!view.filter||active.every(([column,text])=>{const c=columns.indexOf(column);return c<0||display(r,c).toLowerCase().includes(text.trim().toLowerCase());});
 const focusRow=rows[selected[0]],anchorRow=rows[anchor[0]],rangeEnd=anchorRow&&columns[anchor[1]]&&(anchor[0]!==selected[0]||anchor[1]!==selected[1])?{row:sheetRowIdentity(anchorRow),column:columns[anchor[1]],label:cellName(anchor[0],anchor[1])}:undefined;
 const peers=useOfficePresence(item.doc,'cell',focusRow&&columns[selected[1]]?{kind:'cell',sheet:selectedSheet.id,row:sheetRowIdentity(focusRow),column:columns[selected[1]],label:cellName(selected[0],selected[1]),...(rangeEnd?{end:rangeEnd}:{})}:undefined,!!workspace.sync);
 const peerAt=(r:number,c:number)=>peers.find(peer=>peer.focus?.kind==='cell'&&peer.focus.sheet===selectedSheet.id&&peer.focus.row===sheetRowIdentity(rows[r])&&peer.focus.column===columns[c]);
 // Collaborators' selected ranges on this sheet, resolved from stable row/column identities to current indexes.
 const rowIndex=new Map(rows.map((row,r)=>[sheetRowIdentity(row),r])),peerRanges=peers.flatMap(peer=>{const focus=peer.focus;if(focus?.kind!=='cell'||focus.sheet!==selectedSheet.id||!focus.end)return [];const r1=rowIndex.get(focus.row),r2=rowIndex.get(focus.end.row),c1=columns.indexOf(focus.column),c2=columns.indexOf(focus.end.column);if(r1===undefined||r2===undefined||c1<0||c2<0)return [];return [{peer,top:Math.min(r1,r2),bottom:Math.max(r1,r2),left:Math.min(c1,c2),right:Math.max(c1,c2)}];});
 const peerRangeAt=(r:number,c:number)=>peerRanges.find(value=>r>=value.top&&r<=value.bottom&&c>=value.left&&c<=value.right);
 // Unresolved cell comments on this sheet show an Excel-style corner mark that opens the thread.
 const cellComments=new Map<string,{thread:string;count:number}>();if(item.doc.share.has('officeComments'))for(const thread of item.doc.getMap<OfficeCommentThread>('officeComments').values()){if(!validOfficeComment(thread)||thread.resolved||thread.anchor?.kind!=='cell'||thread.anchor.sheet!==selectedSheet.id)continue;const r=rowIndex.get(thread.anchor.row),c=columns.indexOf(thread.anchor.column);if(r===undefined||c<0)continue;const key=r+':'+c,old=cellComments.get(key);cellComments.set(key,{thread:old?.thread??thread.id,count:(old?.count??0)+1});}
 // Focus right away when the cell exists so fast typing cannot land in the previous cell; otherwise wait for the render.
 const focusCell=(r:number,c:number)=>{const run=()=>{const input=document.querySelector<HTMLInputElement>(`.office-spreadsheet-editor input[aria-label="第 ${r+1} 行第 ${c+1} 列"]`);if(!input)return false;input.focus();input.scrollIntoView({block:'nearest',inline:'nearest'});return true;};if(!run())requestAnimationFrame(()=>void run());};
 const locate=(sheet:string,r:number,c:number)=>{setSheetId(sheet);select(r,c);focusCell(r,c);};
 const locatePeer=(peer:typeof peers[number])=>{if(peer.focus?.kind!=='cell')return;const focus=peer.focus,sheet=sheets.find(value=>value.id===focus.sheet);if(!sheet)return;const r=sheet.sheet.toArray().findIndex(row=>sheetRowIdentity(row)===focus.row),c=sheetColumnIds(sheet.sheet).indexOf(focus.column);if(r>=0&&c>=0)locate(sheet.id,r,c);};
 const shell=useRef<HTMLElement>(null),follow=useFollowPeer(peers,locatePeer,shell);
 // Only refresh the formula bar when the active cell's snapshot actually changes.
 const onHandle=(next:SheetCellHandle)=>setHandle(current=>current&&current.id===next.id&&current.value===next.value&&current.kind===next.kind&&current.dirty===next.dirty&&current.busy===next.busy?current:next);
 const run=async(operation:SheetOperation,done?:string)=>{if(disabled)return false;setBusy(true);setOperationError('');setNotice('');try{const result=await workspace.submitSpreadsheetOperation(item,operation);if(result.queued)setOperationError('网络不可用，已保存为离线操作，联网后自动提交');else if(done)setNotice(done);return true;}catch(e){setOperationError(e instanceof Error?e.message:'表格操作失败');await workspace.refreshSpreadsheet(item).catch(()=>{});return false;}finally{setBusy(false);}};
 // Ribbon buttons blur the active cell first; give its commit a moment before requiring a clean grid.
 const settled=async()=>{for(let i=0;i<20&&item.drafts.size;i++)await new Promise(resolve=>setTimeout(resolve,100));if(item.drafts.size)throw new Error('请先提交或处理单元格草稿，再执行此操作');};
 const edit=async(edits:{row:number;column:number;value:string;cellType?:'text'|'auto'|'formula'}[],done:string)=>{try{await settled();if(!edits.length)throw new Error('选区内没有需要修改的单元格');if(edits.length>500)throw new Error('一次最多修改 500 个单元格，请缩小选区');await run({operationId:crypto.randomUUID(),type:'setCells',base:'',cells:cellEdits(item.sheet,edits)},done);}catch(e){setOperationError(e instanceof Error?e.message:'表格操作失败');}};
 const cells=()=>{const list:[number,number][]=[];for(let r=range.top;r<=range.bottom&&r<rows.length;r++)for(let c=range.left;c<=range.right&&c<width;c++)list.push([r,c]);return list;};
 const format=(patch:SheetFormatPatch,target=range)=>{const list:[number,number][]=[];for(let r=target.top;r<=target.bottom&&r<rows.length;r++)for(let c=target.left;c<=target.right&&c<width;c++)list.push([r,c]);if(list.length>SHEET_FORMAT_LIMITS.batch){setOperationError('一次最多设置 2000 个单元格的格式，请缩小选区');return;}void run({operationId:crypto.randomUUID(),type:'setFormat',targets:list.map(([r,c])=>({rowKey:sheetRowIdentity(rows[r]),columnKey:columns[c]})),format:patch,base:''});};
 const structure=async(type:'insertRow'|'deleteRow'|'insertColumn'|'deleteColumn'|'appendRow'|'deleteLastRow',after=false)=>{try{await settled();}catch(e){setOperationError('请先提交或处理单元格草稿，再修改行结构');return;}const row=Math.min(selected[0],Math.max(0,rows.length-1))+(after&&type==='insertRow'?1:0),column=Math.min(selected[1],width-1)+(after&&type==='insertColumn'?1:0);void run({operationId:crypto.randomUUID(),type,row,column,base:structureRevision(item.doc)});};
 const actions:SpreadsheetActions={
  format,error:setOperationError,structure:(type,after)=>void structure(type,after),
  merge:()=>{if(cells().some(([r,c])=>(r!==range.top||c!==range.left)&&item.drafts.has(cellDraftKey(item.sheet,r,c)))){setOperationError('选区内还有未提交的单元格草稿，请先保存或放弃');return;}void run({operationId:crypto.randomUUID(),type:'mergeCells',rowKey:sheetRowIdentity(rows[range.top]),columnKey:columns[range.left],rowSpan:range.bottom-range.top+1,colSpan:range.right-range.left+1,base:''});},
  unmerge:()=>{const owner=merges.owner.get(range.top+':'+range.left),merge=owner?merges.anchors.get(owner):undefined;if(merge)void run({operationId:crypto.randomUUID(),type:'unmergeCells',rowKey:sheetRowIdentity(rows[merge.row]),columnKey:columns[merge.column],base:''});},
  autoSum:fn=>{const plan=autoSumPlan(grid,range,fn);if(!plan.length){setOperationError('未找到可计算的相邻数字，请选中数字区域或其下方的空单元格');return;}if(plan.some(target=>grid[target.row]?.[target.column]))
   {setOperationError('计算结果的目标单元格已有内容，请选中下方的空单元格');return;}void edit(plan.map(target=>({row:target.row,column:target.column,value:target.formula,cellType:'formula' as const})),'已插入 '+plan.map(target=>target.formula).join('、')).then(()=>{const last=plan[plan.length-1];select(last.row,last.column);});},
  sort:descending=>{if(range.bottom===range.top){setOperationError('请先选中要排序的多行区域，可单击列标或在名称框输入区域');return;}if(cells().some(([r,c])=>isFormula(r,c)||merges.owner.has(r+':'+c))){setOperationError('排序区域包含公式或合并单元格，为避免引用错位，请先处理后再排序');return;}
   const order=sortedRows(grid,range,selected[1],descending),changes:{row:number;column:number;value:string;cellType:'text'|'auto'|'formula'}[]=[];order.forEach((source,i)=>{const target=range.top+i;for(let c=range.left;c<=range.right;c++)if(grid[source][c]!==grid[target][c])changes.push({row:target,column:c,value:grid[source][c],cellType:spreadsheetCellType(item.doc,source,c,selectedSheet.id)});});
   if(!changes.length){setNotice('选区已按此顺序排列');return;}void edit(changes,`已按 ${columnName(selected[1])} 列${descending?'降序':'升序'}排序；单元格格式不随行移动`);},
  clear:()=>void edit(cells().filter(([r,c])=>grid[r][c]&&!(merges.owner.has(r+':'+c)&&merges.owner.get(r+':'+c)!==r+':'+c)).map(([r,c])=>({row:r,column:c,value:''})),'已清除内容'),
  copy:()=>{const tsv=rangeToTsv(rows.map((_,r)=>Array.from({length:width},(_,c)=>display(r,c))),range);void navigator.clipboard?.writeText(tsv).then(()=>setNotice(`已复制 ${cells().length} 个单元格`),()=>setOperationError('浏览器不允许写入剪贴板，请使用 Ctrl+C'));},
  paste:()=>setPasteOpen(true),find:()=>{setFind(value=>value??{query:'',replacement:'',index:-1});requestAnimationFrame(()=>findInput.current?.focus());},
  insertFunction:fn=>{if(!handle)return;handle.stage('='+fn+'(',handle.kind==='text'?'formula':handle.kind);handle.focus();},
  comment:()=>[...document.querySelectorAll<HTMLButtonElement>('.office-spreadsheet-editor .office-title-actions button')].find(button=>button.textContent?.startsWith('评论'))?.click(),
  painter:()=>{if(painter){setPainter(undefined);return;}setPainter(sheetCellFormat(item.doc,selectedSheet.id,item.sheet,selected[0],selected[1])??{});setNotice('格式刷：单击或拖选目标单元格');},
 };
 const applyPainter=(target:SheetRange)=>{if(!painter)return;const patch:SheetFormatPatch={bold:painter.bold??null,italic:painter.italic??null,underline:painter.underline??null,color:painter.color??null,fill:painter.fill??null,align:painter.align??null,numberFormat:painter.numberFormat??null};setPainter(undefined);format(patch,target);};
 useEffect(()=>{const up=()=>{dragging.current=false;if(!painter)return;const {anchor:a,selected:s}=selection.current;applyPainter({top:Math.min(a[0],s[0]),left:Math.min(a[1],s[1]),bottom:Math.max(a[0],s[0]),right:Math.max(a[1],s[1])});};document.addEventListener('mouseup',up);return()=>document.removeEventListener('mouseup',up);});
 /** Keyboard like Excel: Enter/Tab/arrows move, Shift extends, Delete clears the selection when not typing. */
 const move=(r:number,c:number,dr:number,dc:number,extend:boolean)=>{let row=r+dr,column=c+dc;while(row>0&&row<rows.length-1&&!visible(row))row+=dr||1;row=Math.max(0,Math.min(rows.length-1,row));column=Math.max(0,Math.min(width-1,column));
  while(hidden(row,column)&&(dr||dc)){const next=[row+dr,column+dc];if(next[0]<0||next[0]>=rows.length||next[1]<0||next[1]>=width)break;[row,column]=next;}
  if(extend){keepAnchor.current=true;setSelected([row,column]);setTimeout(()=>{keepAnchor.current=false;},200);}else select(row,column);focusCell(row,column);};
 const onKey=(r:number,c:number)=>(key:SheetCellKey)=>{
  const edge=(side:'start'|'end')=>!key.editing||key.whole||(side==='start'?key.atStart:key.atEnd);
  if(key.key==='Enter'){move(r,c,key.shift?-1:1,0,false);if(r===(key.shift?0:rows.length-1))(document.activeElement as HTMLElement|null)?.blur();return true;}
  if(key.key==='Tab'){move(r,c,0,key.shift?-1:1,false);return true;}
  if(key.key==='ArrowUp'||key.key==='ArrowDown'){move(r,c,key.key==='ArrowUp'?-1:1,0,key.shift);return true;}
  if((key.key==='ArrowLeft'&&edge('start'))||(key.key==='ArrowRight'&&edge('end'))){move(r,c,0,key.key==='ArrowLeft'?-1:1,key.shift);return true;}
  // Ready mode (nothing typed yet) clears like Excel; once typing, Delete/Backspace edit characters.
  if((key.key==='Delete'||key.key==='Backspace')&&!key.editing&&!readOnly){actions.clear();return true;}
  if(key.mod&&['b','i','u'].includes(key.key.toLowerCase())){const mark=({b:'bold',i:'italic',u:'underline'} as const)[key.key.toLowerCase() as 'b'|'i'|'u'],on=cells().every(([row,col])=>sheetCellFormat(item.doc,selectedSheet.id,item.sheet,row,col)?.[mark]);format({[mark]:on?null:true});return true;}
  if(key.mod&&key.key.toLowerCase()==='c'&&(range.bottom>range.top||range.right>range.left)){actions.copy();return true;}
  if(key.mod&&key.key.toLowerCase()==='f'){actions.find();return true;}
  return false;
 };
 const pick=(top:number,left:number,bottom:number,right:number)=>{setAnchor([top,left]);setSelected([bottom,right]);keepAnchor.current=true;focusCell(top,left);setTimeout(()=>{keepAnchor.current=false;},200);};
 const findHits=find?findMatches(rows.map((_,r)=>Array.from({length:width},(_,c)=>display(r,c))),find.query):[];
 const findNext=()=>{if(!find||!findHits.length){setOperationError(find?.query?'未找到匹配的单元格':'请输入查找内容');return;}setOperationError('');const after=findHits.findIndex(([r,c])=>r>selected[0]||r===selected[0]&&c>selected[1]),index=after<0?0:after,[r,c]=findHits[index];setFind({...find,index});select(r,c);focusCell(r,c);};
 const replace=(all:boolean)=>{if(!find?.query)return;const plan=replacePlan(grid,find.query,find.replacement,(r,c)=>merges.owner.has(r+':'+c)&&merges.owner.get(r+':'+c)!==r+':'+c),edits=all?plan:plan.filter(edit=>edit.row===selected[0]&&edit.column===selected[1]);if(!edits.length){if(all)setOperationError('没有可替换的内容（公式单元格不会被替换）');else findNext();return;}void edit(edits,`已替换 ${edits.length} 个单元格`);};
 const values:string[]=[];for(const [r,c]of cells().slice(0,20000))if(visible(r))values.push(shown(r,c));const stats=selectionStats(values),multi=range.bottom>range.top||range.right>range.left,format0=sheetCellFormat(item.doc,selectedSheet.id,item.sheet,selected[0],selected[1]);
 const status=readOnly?'只读':workspace.sync?'多人协同表格':'本地表格',sourceFormat=workspace.officeResources.get(item.id)?.sourceFormat;
 return <section ref={shell} className="office-spreadsheet-editor office-document-shell" aria-label={name}>
  <OfficeTitleBar kind="sheet" name={name} status={`${status} · ${rows.length} 行 × ${width} 列`}>
   <OfficeActivityButton doc={item.doc} workspace={workspace} id={item.id}/>
   <OfficeComments doc={item.doc} workspace={workspace} id={item.id} captureAnchor={()=>{const row=rows[selected[0]],column=columns[selected[1]];if(row&&column)return {kind:'cell',sheet:selectedSheet.id,row:sheetRowIdentity(row),column,label:cellName(selected[0],selected[1])};}} onLocate={anchor=>{const target=officeAnchorTarget(item.doc,anchor);if(target?.kind==='cell')locate(target.sheet,target.row,target.column);}}/>
   <OfficeMenu label="导出" icon={Download}>{close=><><button role="menuitem" onClick={()=>{close();setDelimitedOpen(true);}}>导出 {sourceFormat==='tsv'?'TSV':'CSV'}</button><button role="menuitem" onClick={()=>{close();setExportKind('text');}}>导出 XLSX 文本副本</button><button role="menuitem" onClick={()=>{close();setExportKind('formula');}}>导出 XLSX 公式副本</button><button role="menuitem" onClick={()=>{close();window.dispatchEvent(new CustomEvent('office-file-action',{detail:{workspace:workspace.id,kind:'exports',id:item.id}}));}}>导出记录</button></>}</OfficeMenu>
  </OfficeTitleBar>
  <SpreadsheetRibbon item={item} range={range} merges={merges} handle={handle} disabled={disabled} readOnly={readOnly} actions={actions} painter={!!painter} view={view} onView={next=>{setView(next);if(!next.filter)setFilters({});}} rowCount={rows.length} columnCount={width}/>
  <div className="office-message-bar">
   {validXlsxSource(source)&&<details className="sheet-import-report" role="note" aria-label="XLSX 导入范围"><summary>XLSX 编辑副本：当前工作表「{selectedSheet.name}」，已接入 {sheets.length} 张，原文件共 {source.sheetCount} 张</summary>已接入加粗/斜体/下划线、RGB 文字与填充色、水平对齐、数字与日期格式及合并单元格；主题色、边框、字体、条件格式、图片、图表、超链接和隐藏状态未接入。公式使用当前基础计算器，不使用原文件缓存冒充重算。{source.textConversions&&'布尔和错误单元格按文本保留。'}原件保持不变。</details>}
   {operationError&&<p role="alert">{operationError}</p>}{notice&&!operationError&&<p role="status">{notice}</p>}
   <SpreadsheetDetachedDrafts item={item}/><SpreadsheetPendingOperations item={item} workspace={workspace} readOnly={readOnly}/>
   <OfficePresenceList peers={peers} describe={peer=>peer.focus?.kind==='cell'?(sheets.find(sheet=>sheet.id===(peer.focus as {sheet:string}).sheet)?.name??'')+' · '+peer.focus.label:undefined} onLocate={locatePeer} following={follow.following} onFollow={follow.follow}/>
  </div>
  <div className="spreadsheet-formula-bar">
   <input className="spreadsheet-name-box" aria-label="名称框" value={nameDraft??(cellName(range.top,range.left)+(multi?':'+cellName(range.bottom,range.right):''))} onFocus={event=>event.currentTarget.select()} onChange={event=>setNameDraft(event.target.value)} onBlur={()=>setNameDraft(undefined)} onKeyDown={event=>{if(event.key==='Escape'){setNameDraft(undefined);event.currentTarget.blur();}if(event.key!=='Enter')return;event.preventDefault();const target=parseCellRange(event.currentTarget.value,rows.length,width);if(!target){setOperationError('请输入如 A1 或 A1:C3 的选区');return;}setOperationError('');setNameDraft(undefined);pick(target.top,target.left,target.bottom,target.right);}}/>
   <button aria-label="取消输入" title="取消输入" disabled={!handle?.dirty} onMouseDown={event=>event.preventDefault()} onClick={()=>{handle?.cancel();handle?.focus();}}><X size={14} aria-hidden="true"/></button>
   <button aria-label="确认输入" title="确认输入" disabled={!handle?.dirty} onMouseDown={event=>event.preventDefault()} onClick={()=>void handle?.commit()}><Check size={14} aria-hidden="true"/></button>
   <label className="spreadsheet-fx" title="插入函数"><span aria-hidden="true">fx</span><select aria-label="函数" value="" disabled={disabled||!handle} onChange={event=>{if(event.target.value)actions.insertFunction(event.target.value);}}><option value="">函数</option>{SHEET_FUNCTIONS.map(([fn,label])=><option key={fn} value={fn}>{fn} · {label}</option>)}</select></label>
   <input aria-label="当前单元格内容" value={handle?.value??''} readOnly={readOnly||!handle||handle.busy} onChange={event=>handle?.stage(event.target.value,handle.kind)} onKeyDown={event=>{if(event.nativeEvent.isComposing)return;if(event.key==='Enter'){event.preventDefault();handle?.focus();}if(event.key==='Escape'){event.preventDefault();handle?.cancel();handle?.focus();}}} onBlur={()=>{if(handle?.dirty)void handle.commit();}}/>
  </div>
  {find&&<div className="spreadsheet-find" role="search" aria-label="查找和替换"><label>查找<input ref={findInput} aria-label="查找内容" value={find.query} onChange={event=>setFind({...find,query:event.target.value,index:-1})} onKeyDown={event=>{if(event.key==='Enter'){event.preventDefault();findNext();}if(event.key==='Escape')setFind(undefined);}}/></label><label>替换为<input aria-label="替换为" value={find.replacement} onChange={event=>setFind({...find,replacement:event.target.value})}/></label><span role="status">{find.query?`${findHits.length} 个匹配`:''}</span><button onClick={findNext}>查找下一个</button><button disabled={disabled||!find.query} onClick={()=>replace(false)}>替换</button><button disabled={disabled||!find.query} onClick={()=>replace(true)}>全部替换</button><button aria-label="关闭查找" onClick={()=>setFind(undefined)}><X size={14} aria-hidden="true"/></button></div>}
  <div className="spreadsheet-scroll" data-gridlines={view.gridlines} data-freeze-row={view.freezeRow} data-freeze-column={view.freezeColumn} data-painter={!!painter||undefined}><table style={{zoom:zoom/100}}><thead><tr><th className="spreadsheet-corner" aria-label="全选" onClick={()=>pick(0,0,rows.length-1,width-1)}/>{Array.from({length:width},(_,c)=><th scope="col" key={c} className={c>=range.left&&c<=range.right?'active':undefined} onClick={()=>pick(0,c,rows.length-1,c)}>{columnName(c)}</th>)}</tr>
   {view.filter&&<tr className="spreadsheet-filter-row"><th aria-label="筛选"/>{columns.map((column,c)=><th key={column}><input aria-label={`筛选 ${columnName(c)} 列`} placeholder="筛选" value={filters[column]??''} onChange={event=>setFilters({...filters,[column]:event.target.value})}/></th>)}</tr>}</thead>
   <tbody>{rows.map((row,r)=>visible(r)&&<tr key={selectedSheet.id+':'+sheetRowIdentity(row)}><th scope="row" className={r>=range.top&&r<=range.bottom?'active':undefined} onClick={()=>pick(r,0,r,width-1)}>{r+1}</th>{Array.from({length:width},(_,c)=>{if(hidden(r,c))return null;const merge=merges.anchors.get(r+':'+c),format=sheetCellFormat(item.doc,selectedSheet.id,item.sheet,r,c),current=selected[0]===r&&selected[1]===c,peer=peerAt(r,c),ranged=peer?undefined:peerRangeAt(r,c),comment=cellComments.get(r+':'+c);
    return <td key={columns[c]} colSpan={merge&&merge.colSpan>1?merge.colSpan:undefined} rowSpan={merge&&merge.rowSpan>1?merge.rowSpan:undefined} className={[current?'selected':'',inRange(r,c)&&!current?'in-range':''].filter(Boolean).join(' ')} style={{background:format?.fill?'#'+format.fill:undefined,...(peer?{outline:'2px solid '+peer.color,outlineOffset:-2}:ranged?{boxShadow:'inset 0 0 0 999px '+ranged.peer.color+'1f'}:{})}} title={peer?peer.name+' 正在此单元格':ranged?ranged.peer.name+' 已选中此区域':undefined} data-peer={peer?.name} data-peer-range={ranged?.peer.name}
     onMouseDown={event=>{if(event.shiftKey){event.preventDefault();setSelected([r,c]);return;}dragging.current=true;}} onMouseEnter={event=>{if(dragging.current&&event.buttons===1&&(selected[0]!==r||selected[1]!==c))setSelected([r,c]);}}>
     <SpreadsheetCell item={item} workspace={workspace} row={row} column={c} format={format} active={current} readOnly={disabled} onHandle={onHandle} onKey={onKey(r,c)} onFocus={()=>{if(keepAnchor.current){keepAnchor.current=false;return;}if(selected[0]!==r||selected[1]!==c||anchor[0]!==r||anchor[1]!==c)select(r,c);}}/>
     {comment&&<button type="button" tabIndex={-1} className="spreadsheet-comment-mark" aria-label={`查看 ${cellName(r,c)} 的批注`+(comment.count>1?`（${comment.count} 条）`:'')} title="查看批注" onMouseDown={event=>{event.preventDefault();event.stopPropagation();}} onClick={()=>requestOfficeThread(workspace.id,item.id,comment.thread)}/>}
     {!view.formulas&&isFormula(r,c)&&<small className="spreadsheet-result">{formatSheetValue(String(results[sheetName]?.[r]?.[c]??''),format?.numberFormat)}</small>}</td>;})}</tr>)}</tbody></table></div>
  <div className="spreadsheet-sheet-bar"><nav aria-label="工作表列表">{sheets.map(sheet=><button disabled={busy} key={sheet.id} aria-current={selectedSheet.id===sheet.id?'page':undefined} onClick={()=>{select(0,0);setSheetId(sheet.id);}}>{sheet.name}</button>)}</nav><WorksheetControls item={item} workspace={workspace} busy={busy} onBusy={setBusy} onSelect={setSheetId}/></div>
  <OfficeStatusBar><span>{handle?.dirty?'编辑':'就绪'}</span>{view.filter&&active.length>0&&<span>已筛选：显示 {rows.filter((_,r)=>r>0&&visible(r)).length} / {Math.max(0,rows.length-1)} 行</span>}{format0?.numberFormat&&<span>格式 {format0.numberFormat}</span>}
   <span className="spreadsheet-stats" aria-label="选区统计">{multi&&stats.average!==undefined&&<span>平均值 {formatStat(stats.average)}</span>}{multi&&<span>计数 {stats.count}</span>}{stats.numeric>0&&multi&&<span>求和 {formatStat(stats.sum)}</span>}</span>
   <label>缩放 <select aria-label="表格缩放" value={zoom} onChange={event=>setZoom(Number(event.target.value))}>{[50,75,90,100,125,150,200].map(value=><option key={value} value={value}>{value}%</option>)}</select></label></OfficeStatusBar>
  {pasteOpen&&<SpreadsheetPasteDialog item={item} workspace={workspace} start={selected} onClose={()=>setPasteOpen(false)}/>}
  {delimitedOpen&&<DelimitedExportDialog workspace={workspace} id={item.id} name={name} sheetId={selectedSheet.id} sheetName={selectedSheet.name} initialFormat={sourceFormat==='tsv'?'tsv':'csv'} onClose={()=>setDelimitedOpen(false)}/>}
  {exportKind&&<SpreadsheetExportDialog workspace={workspace} id={item.id} name={name} kind={exportKind} onClose={()=>setExportKind(undefined)}/>}
 </section>;
}
