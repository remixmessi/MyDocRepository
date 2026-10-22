import {useState} from 'react';
import {Modal} from './dialogs';
import type {LocalWorkspace} from './workspace';
import {cloudRequest} from '../../../common/nbstore/src/cloud';
import {parseDelimited} from '../../../common/nbstore/src/office';
import type {DelimitedFormat,DelimitedTextMode} from '../../../common/nbstore/src/office-delimited-export';
import {downloadOfficeBytes} from './office-download';

type Preview={bytes:Uint8Array;rows:string[][];rowCount:number;columns:number};
export function DelimitedExportDialog({workspace,id,name,initialFormat,onClose,sheetId,sheetName}:{workspace:LocalWorkspace;id:string;name:string;initialFormat:DelimitedFormat;sheetId?:string;sheetName?:string;onClose:()=>void}){
 const [content,setContent]=useState<'text'|'values'>('values');const [format,setFormat]=useState(initialFormat),[mode,setMode]=useState<DelimitedTextMode>('escaped'),[busy,setBusy]=useState(false),[error,setError]=useState(''),[preview,setPreview]=useState<Preview>();
 const generate=async()=>{if(busy)return;setBusy(true);setError('');setPreview(undefined);try{
  await workspace.prepareOfficeBackup(id);
  const response=await cloudRequest('/office-export?'+new URLSearchParams({workspace:workspace.id,id,format,scope:mode+'-'+content,sheet:sheetId??'first'}));
  const bytes=new Uint8Array(await response.arrayBuffer());if(bytes.length>16*1024*1024)throw new Error('导出文件超过 16 MB');
  const values=parseDelimited(new TextDecoder('utf-8',{fatal:true}).decode(bytes),format==='csv'?',':'\t');
  setPreview({bytes,rows:values.slice(0,5).map(row=>row.slice(0,5)),rowCount:values.length,columns:Math.max(0,...values.map(row=>row.length))});
 }catch(error){setError(error instanceof Error?error.message:'生成预览失败，请重试');}finally{setBusy(false);}};
 return <Modal title="导出 CSV/TSV" onClose={()=>{if(!busy)onClose();}} className="delimited-export-dialog"><p>{sheetName&&<>工作表「{sheetName}」。</>}仅导出当前单工作表的已保存文本，不包含其他工作表、样式、图表、合并、批注或历史。{content==='text'?'公式保留原文，不计算结果。':'公式结果由服务端从固定快照计算；错误、循环或不支持的公式会阻止导出，可选择公式原文。'}</p>
  <label className="form-field">导出内容<select aria-label="导出内容" value={content} disabled={busy} onChange={event=>{setContent(event.target.value as 'text'|'values');setPreview(undefined);setError('');}}><option value="values">服务端计算结果</option><option value="text">公式原文（不计算）</option></select></label>
  <label className="form-field">文件格式<select aria-label="文件格式" value={format} disabled={busy} onChange={event=>{setFormat(event.target.value as DelimitedFormat);setPreview(undefined);setError('');}}><option value="csv">CSV（逗号分隔）</option><option value="tsv">TSV（制表符分隔，CSV 式引号）</option></select></label>
  <label className="form-field">公式样式文本<select aria-label="公式样式文本" value={mode} disabled={busy} onChange={event=>{setMode(event.target.value as DelimitedTextMode);setPreview(undefined);setError('');}}><option value="escaped">添加单引号前缀</option><option value="raw">保留原文</option></select></label>
  <p>{mode==='escaped'?'对以 =、+、-、@ 等开头的字段（含前置空白），以及以制表符或换行开头的字段添加单引号。这会改变字段内容，预览将展示修改后的值。':'不修改字段内容；外部表格软件可能将公式样式文本作为公式执行。'}</p>
  <p>编码为 UTF-8，带 BOM；记录之间使用 CRLF，字段内换行保留。引号用于字段转义，不能阻止公式执行。不同软件处理单引号的方式不同；需要明确文本类型时，请使用 XLSX 文本副本。</p>
  {error&&<p role="alert" className="form-error">{error}</p>}
  {preview&&<section aria-label="导出内容预览"><p role="status">{preview.rowCount} 行 × {preview.columns} 列 · 最多预览前 5 行、5 列，每格 120 字符</p>{preview.rows.length?<div className="delimited-export-preview"><table><thead><tr>{Array.from({length:Math.min(5,preview.columns)},(_,c)=><th scope="col" key={c}>第 {c+1} 列</th>)}</tr></thead><tbody>{preview.rows.map((row,r)=><tr key={r}>{row.map((value,c)=><td key={c}>{value.slice(0,120)}{value.length>120?'…':''}</td>)}</tr>)}</tbody></table></div>:<p>工作表为空，导出文件仅包含编码标记。</p>}<p>下载内容已固定。预览生成后的编辑不会进入本次下载，可重新生成预览获取最新内容。</p></section>}
  <div className="dialog-actions"><button disabled={busy} onClick={onClose}>取消</button><button disabled={busy} onClick={()=>void generate()}>{busy?'正在生成…':preview?'重新生成预览':'生成预览'}</button><button className="primary" disabled={busy||!preview} onClick={()=>{if(!preview)return;downloadOfficeBytes(name.replace(/\.[^.]+$/,'')+(content==='values'?'-计算结果.':'-文本副本.')+format,preview.bytes,format==='csv'?'text/csv;charset=utf-8':'text/tab-separated-values;charset=utf-8');onClose();}}>下载文件</button></div>
 </Modal>;
}
