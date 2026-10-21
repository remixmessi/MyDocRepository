import {createPortal} from 'react-dom';
import {useEffect,useMemo,useState} from 'react';
import type {LocalWorkspace} from './workspace';
import {Modal} from './dialogs';
import {prepareDelimitedImport,type DelimitedImportOptions} from '../../../common/nbstore/src/office-delimited-import';
import type {TextEncoding} from '../../../common/nbstore/src/office';
const ACCEPT='.docx,.docm,.dotx,.dotm,.xlsx,.xlsm,.xltx,.xltm,.csv,.tsv,.pdf,.pptx,.pptm,.potx,.potm';
export function OfficeImportDialog({workspace,onClose,onFinish}:{workspace:LocalWorkspace;onClose:()=>void;onFinish:(id:string)=>void}){
 const [file,setFile]=useState<File>(),[bytes,setBytes]=useState<Uint8Array>(),[loading,setLoading]=useState(false),[readError,setReadError]=useState(''),[busy,setBusy]=useState(false),[error,setError]=useState('');
 const [encoding,setEncoding]=useState<TextEncoding|''>(''),[delimiter,setDelimiter]=useState<','|'\t'>(','),[quoteMode,setQuoteMode]=useState<'csv'|'none'>('csv');
 const delimited=!!file&&/\.(csv|tsv)$/i.test(file.name);
 const options:DelimitedImportOptions={...(encoding?{encoding}:{}),delimiter,quoteMode};
 useEffect(()=>{let active=true;setBytes(undefined);setReadError('');setLoading(false);if(!file)return;if(file.size>20*1024*1024){setReadError('原件超过 20 MB');return;}if(!delimited)return;setLoading(true);void file.arrayBuffer().then(data=>{if(active)setBytes(new Uint8Array(data));}).catch(()=>{if(active)setReadError('文件读取失败，请重新选择');}).finally(()=>{if(active)setLoading(false);});return()=>{active=false;};},[file,delimited]);
 const parsed=useMemo(()=>{if(!bytes)return undefined;try{return {value:prepareDelimitedImport(bytes,{...(encoding?{encoding}:{}),delimiter,quoteMode}),error:''};}catch(error){return {value:undefined,error:error instanceof Error?error.message:'无法解析 CSV/TSV'};}},[bytes,encoding,delimiter,quoteMode]);
 const ready=!!file&&file.size<=20*1024*1024&&!loading&&!readError&&(!delimited||!!parsed?.value)&&!workspace.readOnly;
 return createPortal(<Modal title="导入办公文件" className="delimited-export-dialog" onClose={()=>{if(!busy)onClose();}}><form onSubmit={event=>{event.preventDefault();if(!ready||!file||busy)return;setBusy(true);setError('');void workspace.importOfficeFile(file,delimited?options:undefined).then(onFinish).catch(error=>setError(error instanceof Error?error.message:'导入失败，请重试')).finally(()=>setBusy(false));}}>
  <p>原件按原始字节保留，最多 20 MB。导入将创建新文件。</p>
  <label className="form-field">办公文件<input aria-label="办公文件" type="file" accept={ACCEPT} disabled={busy||workspace.readOnly} onChange={event=>{const selected=event.target.files?.[0];setFile(selected);setBytes(undefined);setError('');setReadError('');setEncoding('');setDelimiter(selected&&/\.tsv$/i.test(selected.name)?'\t':',');setQuoteMode('csv');}}/></label>
  {file&&<p className="office-backup-filename">{file.name} · {(file.size/1024).toFixed(1)} KB</p>}
  {delimited&&<>
   <label className="form-field">文本编码<select aria-label="导入文本编码" value={encoding} disabled={busy} onChange={event=>{setEncoding(event.target.value as TextEncoding|'');setError('');}}><option value="">自动（BOM 优先，否则 UTF-8）</option><option value="utf-8">UTF-8</option><option value="utf-16le">UTF-16LE</option><option value="utf-16be">UTF-16BE</option></select></label>
   <label className="form-field">分隔符<select aria-label="导入分隔符" value={delimiter} disabled={busy} onChange={event=>{const value=event.target.value as ','|'\t';setDelimiter(value);if(value===',')setQuoteMode('csv');setError('');}}><option value=",">逗号</option><option value={'\t'}>制表符</option></select></label>
   <label className="form-field">引号模式<select aria-label="导入引号模式" value={quoteMode} disabled={busy} onChange={event=>{setQuoteMode(event.target.value as 'csv'|'none');setError('');}}><option value="csv">CSV 式双引号转义</option><option value="none" disabled={delimiter!== '\t'}>无引号转义（仅 TSV）</option></select></label>
   <p>首行保留为数据，不自动移除表头。前导零、长编号和公式样式文本均保留原文；所有导入单元格默认为文本，不计算 = 开头内容；导入后可为选中单元格选择公式类型。尚不支持日期等列类型转换。</p>
   {loading&&<p role="status">正在读取并解析文件…</p>}
   {parsed?.value&&<section aria-label="导入内容预览"><p role="status">{parsed.value.rows.length} 行 · {parsed.value.settings.encoding.toUpperCase()} · {parsed.value.settings.bom?'带 BOM':'无 BOM'} · 最多预览前 5 行、5 列，每格 120 字符</p>{parsed.value.rows.length?<div className="delimited-export-preview"><table><thead><tr>{Array.from({length:Math.min(5,Math.max(...parsed.value.rows.map(row=>row.length)))},(_,c)=><th scope="col" key={c}>第 {c+1} 列</th>)}</tr></thead><tbody>{parsed.value.rows.slice(0,5).map((row,r)=><tr key={r}>{row.slice(0,5).map((value,c)=><td key={c}>{value.slice(0,120)}{value.length>120?'…':''}</td>)}</tr>)}</tbody></table></div>:<p>文件为空，将保留原件并创建空工作表。</p>}</section>}
  </>}
  {(readError||parsed?.error||error)&&<p role="alert" className="form-error">{readError||parsed?.error||error}</p>}
  {workspace.readOnly&&<p>当前工作区只读，无法导入。</p>}
  <div className="dialog-actions"><button type="button" onClick={onClose} disabled={busy}>取消</button><button className="primary" disabled={!ready||busy}>{busy?'正在导入…':'导入'}</button></div>
 </form></Modal>,document.body);
}
