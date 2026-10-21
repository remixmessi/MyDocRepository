import {sheetColumnIdentity,sheetRowIdentity} from '../../../common/nbstore/src/sheet-columns';
import {useEffect,useRef,useState} from 'react';
import type * as Y from 'yjs';
import {CloudError,officeDocumentGeneration} from '../../../common/nbstore/src/cloud';
import {cellDraftKey,cellRevision,type SheetOperation} from '../../../common/nbstore/src/spreadsheet-operation';
import {spreadsheetCellType,type SpreadsheetCellType} from '../../../common/nbstore/src/spreadsheet-cell-type';
import type {LocalWorkspace,OfficeSpreadsheetDocument} from './workspace';
import {isNetworkFailure,OFFLINE_DRAFT} from './spreadsheet-operation-queue';
import {spreadsheetDraftPrefix} from './spreadsheet-draft-store';
import {formatSheetValue,isDateNumberFormat,parseSheetDateInput,sheetDateEditText,type SheetCellFormat} from '../../../common/nbstore/src/spreadsheet-format';
/** What the formula bar and ribbon need from the active cell. */
export interface SheetCellHandle {id:string;value:string;kind:SpreadsheetCellType;dirty:boolean;busy:boolean;stage:(value:string,kind:SpreadsheetCellType)=>void;commit:()=>Promise<void>;cancel:()=>void;focus:()=>void}
/** Keys the grid handles (move, extend, clear); `editing` means the user has typed into this cell. */
export type SheetCellKey={key:string;shift:boolean;mod:boolean;editing:boolean;atStart:boolean;atEnd:boolean;whole:boolean};

export function SpreadsheetCell({item,workspace,row,column,readOnly,active,onFocus,format,onHandle,onKey}:{item:OfficeSpreadsheetDocument;workspace:LocalWorkspace;row:Y.Array<string>;column:number;readOnly:boolean;active:boolean;onFocus:()=>void;format?:SheetCellFormat;onHandle?:(handle:SheetCellHandle)=>void;onKey?:(key:SheetCellKey)=>boolean}){
 const [indexKey]=useState(()=>cellDraftKey(item.sheet,item.sheet.toArray().indexOf(row),column));const saved=item.drafts.get(indexKey);
 const currentIndex=()=>item.sheet.toArray().indexOf(row),currentType=()=>spreadsheetCellType(item.doc,currentIndex(),column,item.sheetId);
 const [value,setValue]=useState(()=>saved?.value??row.get(column)??''),[kind,setKind]=useState<SpreadsheetCellType>(()=>saved?.cellType??currentType()),[dirty,setDirty]=useState(!!saved),[busy,setBusy]=useState(false),[error,setError]=useState(saved?.error??''),[conflict,setConflict]=useState(saved?.conflict??false);
 const generation=useRef(saved?.generation??(saved&&!saved.operation?.officeGeneration&&officeDocumentGeneration(workspace.id,item.id)?.startsWith('epoch-')?'legacy-unknown':saved?.operation?.officeGeneration??officeDocumentGeneration(workspace.id,item.id)));
 const [focused,setFocused]=useState(false),date=isDateNumberFormat(format?.numberFormat)&&kind!=='text';
 const input=useRef<HTMLInputElement>(null),valueRef=useRef(value),kindRef=useRef(kind);valueRef.current=value;kindRef.current=kind;
 const dirtyRef=useRef(!!saved),base=useRef(saved?.base??''),pending=useRef<SheetOperation|undefined>(saved?.operation);if(saved?.operation&&!saved.operation.officeGeneration&&officeDocumentGeneration(workspace.id,item.id)?.startsWith('epoch-'))saved.operation.officeGeneration='legacy-unknown';
 useEffect(()=>{const update=()=>{if(!dirtyRef.current){setValue(row.get(column)??'');setKind(currentType());}};row.observe(update);update();return()=>row.unobserve(update);},[row,column]);
 // The workspace replays offline drafts on reconnect; follow its outcome here.
 useEffect(()=>{const key=spreadsheetDraftPrefix(item.id)+indexKey,watch=(event:Y.YMapEvent<unknown>)=>{if(!event.keysChanged.has(key)||busy)return;const draft=workspace.uiPreferences.get(key) as typeof saved;if(!draft){if(dirtyRef.current){dirtyRef.current=false;pending.current=undefined;setDirty(false);setError('');setConflict(false);setValue(row.get(column)??'');setKind(currentType());}}else if(draft.error&&!draft.error.startsWith(OFFLINE_DRAFT)){setError(draft.error);setConflict(!!draft.conflict);}};workspace.uiPreferences.observe(watch);return()=>workspace.uiPreferences.unobserve(watch);},[workspace,item.id,indexKey,busy]);
 const stage=(value:string,cellType:SpreadsheetCellType)=>{
  if(readOnly||busy)return;if(!dirtyRef.current){base.current=cellRevision(item.sheet,currentIndex(),column);generation.current=officeDocumentGeneration(workspace.id,item.id);}dirtyRef.current=true;setDirty(true);setValue(value);setKind(cellType);valueRef.current=value;kindRef.current=cellType;item.drafts.set(indexKey,{generation:generation.current,row:currentIndex(),column,value,cellType,base:base.current});pending.current=undefined;setError('');setConflict(false);
 };
 const commit=async(rebase=false)=>{
  if(!dirtyRef.current||busy||readOnly)return;
  const index=currentIndex();if(index<0){setError('此行已删除，请保留输入内容');return;}
  if(rebase){base.current=cellRevision(item.sheet,index,column);generation.current=officeDocumentGeneration(workspace.id,item.id);pending.current=undefined;}
  // Date-formatted cells store Excel serials; a typed date becomes its serial, anything else stays as typed.
  const value=valueRef.current,kind=kindRef.current,stored=isDateNumberFormat(format?.numberFormat)&&kind!=='text'?parseSheetDateInput(value)??value:value;
  const operation=pending.current??{operationId:crypto.randomUUID(),officeGeneration:generation.current,sheet:item.sheetId,type:'setCell' as const,rowKey:sheetRowIdentity(row),columnKey:sheetColumnIdentity(item.sheet,column),row:index,column,value:stored,cellType:kind,base:base.current};pending.current=operation;item.drafts.set(indexKey,{generation:generation.current,row:index,column,value,cellType:kind,base:base.current,operation});
  setBusy(true);setError('');
  try{await workspace.submitSpreadsheetOperation(item,operation);item.drafts.delete(indexKey);dirtyRef.current=false;setDirty(false);setValue(row.get(column)??'');setKind(currentType());pending.current=undefined;setConflict(false);}
  catch(e){const changed=e instanceof CloudError&&e.status===409;setConflict(changed);const message=isNetworkFailure(e)?OFFLINE_DRAFT+'，联网后自动提交，输入已保存在本机':e instanceof Error?e.message:'提交失败，输入已保留';setError(message);item.drafts.set(indexKey,{generation:generation.current,row:index,column,value,cellType:kind,base:base.current,operation,error:message,conflict:changed});if(changed)await workspace.refreshSpreadsheet(item).catch(()=>{});}
  finally{setBusy(false);}
 };
 const cancel=()=>{item.drafts.delete(indexKey);dirtyRef.current=false;setDirty(false);setValue(row.get(column)??'');setKind(currentType());pending.current=undefined;setError('');setConflict(false);};
 const index=currentIndex(),labels={text:'文本',formula:'公式',auto:'自动'};
 const editText=dirty?value:date?sheetDateEditText(value):value;
 useEffect(()=>{if(active)onHandle?.({id:indexKey,value:editText,kind,dirty,busy,stage,commit:()=>commit(),cancel,focus:()=>requestAnimationFrame(()=>{const element=input.current;if(!element)return;element.focus();element.setSelectionRange(element.value.length,element.value.length);})});});
 const shown=dirty?value:focused?(date?sheetDateEditText(value):value):kind==='text'||value.trim().startsWith('=')?value:formatSheetValue(value,format?.numberFormat);
 const style={fontWeight:format?.bold?700:undefined,fontStyle:format?.italic?'italic':undefined,textDecoration:format?.underline?'underline':undefined,color:format?.color?'#'+format.color:undefined,textAlign:format?.align};
 return <><input ref={input} maxLength={100000} aria-label={`第 ${index+1} 行第 ${column+1} 列`} readOnly={readOnly||busy} value={shown} style={style} onFocus={()=>{setFocused(true);onFocus();if(!dirtyRef.current)base.current=cellRevision(item.sheet,index,column);}}
  onChange={event=>stage(event.target.value,kind)} onBlur={()=>{setFocused(false);if(!error)void commit();}} onKeyDown={event=>{
   if(event.nativeEvent.isComposing)return;
   if(event.key==='Escape'&&dirtyRef.current){event.preventDefault();cancel();return;}
   const element=event.currentTarget,start=element.selectionStart??0,end=element.selectionEnd??0;
   if(onKey?.({key:event.key,shift:event.shiftKey,mod:event.ctrlKey||event.metaKey,editing:dirtyRef.current,atStart:start===0&&end===0,atEnd:start===element.value.length&&end===start,whole:start===0&&end===element.value.length}))event.preventDefault();
   else if(event.key==='Enter'){event.preventDefault();element.blur();}
  }}/>
  {busy&&<small role="status">正在提交…</small>}
  {error&&<div role="alert"><p>{error}</p>{conflict&&<p>最新内容：{row.get(column)||'（空）'}（{labels[currentType()]}）；本人输入：{value||'（空）'}（{labels[kind]}）</p>}
   <button disabled={readOnly||busy} onClick={()=>void commit(conflict)}>{conflict?'重新提交':'重试提交'}</button>
   <button disabled={busy} onClick={()=>{item.drafts.delete(indexKey);dirtyRef.current=false;setDirty(false);setValue(row.get(column)??'');setKind(currentType());pending.current=undefined;setError('');setConflict(false);}}>使用最新内容</button>
  </div>}
 </>;
}
