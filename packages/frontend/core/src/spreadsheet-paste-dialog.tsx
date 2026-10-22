import {useState} from 'react';
import {parseDelimitedGrid} from '../../../common/nbstore/src/office-delimited-import';
import {cellRevision,MAX_SHEET_BATCH_CELLS,type SheetOperation} from '../../../common/nbstore/src/spreadsheet-operation';
import {CloudError} from '../../../common/nbstore/src/cloud';
import {Modal} from './dialogs';
import type {LocalWorkspace,OfficeSpreadsheetDocument} from './workspace';

interface PastePreview {operation:SheetOperation;before:string[];}
export function SpreadsheetPasteDialog({item,workspace,start,onClose}:{item:OfficeSpreadsheetDocument;workspace:LocalWorkspace;start:[number,number];onClose:()=>void}){
 const [text,setText]=useState(''),[row,setRow]=useState(String(start[0]+1)),[column,setColumn]=useState(String(start[1]+1));
 const [preview,setPreview]=useState<PastePreview>(),[busy,setBusy]=useState(false),[error,setError]=useState(''),[conflict,setConflict]=useState(false);
 const readOnly=workspace.documentReadOnly(item.id);
 const clear=()=>{setPreview(undefined);setError('');setConflict(false);};
 const prepare=()=>{
  clear();try{
   if(readOnly)throw new Error('当前文件只读');
   if(item.drafts.size)throw new Error('请先提交或处理单元格草稿，再预览批量粘贴');
   if(!text||text.length>100000)throw new Error('请输入粘贴内容，最多 10 万字符');
   if(!/^[1-9]\d*$/.test(row)||!/^[1-9]\d*$/.test(column))throw new Error('起始行列必须是正整数');
   const r=Number(row)-1,c=Number(column)-1,values=parseDelimitedGrid(text,'\t');
   if(!Number.isSafeInteger(r)||!Number.isSafeInteger(c)||!values.length)throw new Error('粘贴范围无效');
   if(values.reduce((count,values)=>count+values.length,0)>MAX_SHEET_BATCH_CELLS)throw new Error('每次最多粘贴 500 格，请拆分后操作');
   const before:string[]=[],cells=values.flatMap((values,offset)=>values.map((value,index)=>{
    const target=item.sheet.get(r+offset),col=c+index;
    if(!target||col>=target.length)throw new Error('粘贴范围超出现有网格，请调整起始位置或范围');
    if(value.length>10000)throw new Error('单元格内容超过 1 万字符');
    before.push(target.get(col)??'');return {row:r+offset,column:col,value,cellType:'text' as const,base:cellRevision(item.sheet,r+offset,col)};
   }));
   const operation:SheetOperation={operationId:crypto.randomUUID(),type:'setCells',base:'',cells};
   if(new TextEncoder().encode(JSON.stringify(operation)).length>250000)throw new Error('粘贴请求超过大小上限，请拆分后操作');
   setPreview({operation,before});
  }catch(e){setError(e instanceof Error?e.message:'无法生成预览');}
 };
 const submit=async()=>{
  if(!preview||busy||readOnly||conflict)return;
  if(item.drafts.size){setError('请先处理单元格草稿');return;}
  setBusy(true);setError('');
  try{await workspace.submitSpreadsheetOperation(item,preview.operation);onClose();}
  catch(e){setError(e instanceof Error?e.message:'提交失败，粘贴内容已保留');if(e instanceof CloudError&&e.status===409){setConflict(true);await workspace.refreshSpreadsheet(item).catch(()=>{});}}
  finally{setBusy(false);}
 };
 return <Modal className="sheet-paste-dialog" title="批量粘贴" onClose={()=>{if(!busy)onClose();}}>
  <p>从电子表格复制后粘贴到下方，使用制表符分列、换行分行。全部按文本保存，保留前导零及公式原文；空字段会清空目标单元格。每次最多 500 格，整批一起保存。</p>
  <label className="form-field">起始行<input aria-label="粘贴起始行" inputMode="numeric" value={row} disabled={busy||readOnly} onChange={e=>{setRow(e.target.value);clear();}}/></label>
  <label className="form-field">起始列<input aria-label="粘贴起始列" inputMode="numeric" value={column} disabled={busy||readOnly} onChange={e=>{setColumn(e.target.value);clear();}}/></label>
  <label className="form-field">粘贴内容<textarea aria-label="批量粘贴内容" rows={6} maxLength={100001} value={text} disabled={busy||readOnly} onChange={e=>{setText(e.target.value);clear();}}/></label>
  <button disabled={busy||readOnly} onClick={prepare}>{preview?'重新生成预览':'预览粘贴'}</button>
  {preview&&<div aria-label="批量粘贴覆盖预览" className="sheet-paste-preview"><p>将覆盖 {preview.operation.cells!.length} 格（文本类型）。以下显示前 20 格的“旧值 → 新值”。</p><ol>{preview.operation.cells!.slice(0,20).map((cell,index)=><li key={cell.row+':'+cell.column}>第 {cell.row+1} 行第 {cell.column+1} 列：{preview.before[index]||'（空）'} → {cell.value||'（空）'}</li>)}</ol></div>}
  {error&&<p role="alert" className="form-error">{error}{conflict&&'。内容已保留，请重新生成预览并核对最新值后再提交。'}</p>}
  <div className="dialog-actions"><button disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={!preview||busy||readOnly||conflict} onClick={()=>void submit()}>{busy?'正在提交…':'确认整批粘贴'}</button></div>
 </Modal>;
}
