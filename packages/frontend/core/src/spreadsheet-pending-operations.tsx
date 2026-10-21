import {useEffect,useState} from 'react';
import {structureRevision} from '../../../common/nbstore/src/spreadsheet-operation';
import {downloadOfficeBytes} from './office-download';
import {isStructuralSheetOperation,OFFLINE_DRAFT,sheetQueuePrefix} from './spreadsheet-operation-queue';
import {spreadsheetDraftPrefix} from './spreadsheet-draft-store';
import type {LocalWorkspace,OfficeSpreadsheetDocument} from './workspace';
const LABELS:Record<string,string>={insertRow:'插入行',deleteRow:'删除行',insertColumn:'插入列',deleteColumn:'删除列',appendRow:'新增行',deleteLastRow:'删除末行',setCells:'批量粘贴',setFormat:'设置格式',mergeCells:'合并单元格',unmergeCells:'取消合并',addSheet:'新增工作表',deleteSheet:'删除工作表',moveSheetLeft:'左移工作表',moveSheetRight:'右移工作表',renameSheet:'重命名工作表'};
/** Offline operations wait here in order; conflicts stop the queue until the author decides. */
export function SpreadsheetPendingOperations({item,workspace,readOnly}:{item:OfficeSpreadsheetDocument;workspace:LocalWorkspace;readOnly:boolean}){
 const [,refresh]=useState(0),[busy,setBusy]=useState(false);
 useEffect(()=>{const prefixes=[sheetQueuePrefix(item.id),spreadsheetDraftPrefix(item.id)],update=(event:{keysChanged:Set<string>})=>{if([...event.keysChanged].some(key=>prefixes.some(prefix=>key.startsWith(prefix))))refresh(n=>n+1);};workspace.uiPreferences.observe(update);return()=>workspace.uiPreferences.unobserve(update);},[workspace,item.id]);
 const entries=item.queue.entries(),offlineDrafts=[...item.drafts.values()].filter(draft=>draft.operation&&draft.error?.startsWith(OFFLINE_DRAFT)).length,waiting=entries.length+offlineDrafts;
 const drain=async()=>{if(busy||readOnly)return;setBusy(true);try{await workspace.drainSpreadsheetQueue(item);}finally{setBusy(false);refresh(n=>n+1);}};
 // Replay on reconnect, when opening the file and periodically while anything is waiting.
 useEffect(()=>{if(!waiting||readOnly)return;const online=()=>void drain();window.addEventListener('online',online);void drain();const timer=setInterval(()=>void drain(),15000);return()=>{window.removeEventListener('online',online);clearInterval(timer);};},[waiting>0,readOnly]);
 if(!waiting)return null;const blocked=entries.find(([,entry])=>entry.error);
 return <section className="sheet-pending-operations" aria-label="离线待提交操作"><p role="status">{entries.length?`${entries.length} 个表格操作`:''}{entries.length&&offlineDrafts?'、':''}{offlineDrafts?`${offlineDrafts} 个单元格`:''}已保存在本机，联网后按顺序自动提交{blocked?'；以下操作未能提交，后续操作暂停，请处理':''}。</p>
  <ol>{entries.map(([key,entry])=><li key={key}><strong>{LABELS[entry.operation.type]??entry.operation.type}</strong><span>{new Date(entry.created).toLocaleTimeString()}</span>{entry.error?<span role="alert">{entry.error}</span>:<span>等待提交</span>}
   {entry.error&&<>{entry.conflict&&isStructuralSheetOperation(entry.operation)&&<button disabled={readOnly||busy} onClick={()=>{item.queue.update(key,{operation:{...entry.operation,operationId:crypto.randomUUID(),base:structureRevision(item.doc),officeGeneration:undefined},error:undefined,conflict:undefined});void drain();}}>按最新结构重新提交</button>}<button disabled={readOnly||busy} onClick={()=>{item.queue.update(key,{error:undefined,conflict:undefined});void drain();}}>重试</button><button disabled={busy} onClick={()=>item.queue.delete(key)}>放弃此操作</button></>}</li>)}</ol>
  <div className="dialog-actions"><button disabled={readOnly||busy} onClick={()=>void drain()}>{busy?'正在提交…':'立即提交'}</button><button onClick={()=>downloadOfficeBytes('表格离线操作.json',new TextEncoder().encode(JSON.stringify({format:'mydoc-sheet-pending-operations',version:1,resource:item.id,operations:entries.map(([,entry])=>entry)},null,2)),'application/json')}>下载离线操作</button></div>
 </section>;
}
