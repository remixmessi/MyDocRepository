import {createPortal} from 'react-dom';
import {useEffect,useMemo,useRef,useState} from 'react';
import type {LocalWorkspace} from './workspace';
import {Modal} from './dialogs';
import {useFileDrop,FileDropHint} from '../../../editor/blocks/src/file-drop';
import {prepareDelimitedImport,type DelimitedImportOptions} from '../../../common/nbstore/src/office-delimited-import';
import {OFFICE_IMPORT_EXTENSIONS,officeDelimited,officeImportProblem,type TextEncoding} from '../../../common/nbstore/src/office';
const ACCEPT=OFFICE_IMPORT_EXTENSIONS.map(ext=>'.'+ext).join(',');
const MAX_BATCH=20;
/** Imports one file with CSV/TSV options, or a batch of Word / Excel / PPT / PDF files. Dropped files start at once. */
export function OfficeImportDialog({workspace,initialFiles,onClose,onFinish}:{workspace:LocalWorkspace;initialFiles?:File[];onClose:()=>void;onFinish:(ids:string[])=>void}){
 const [files,setFiles]=useState<File[]>(initialFiles??[]),[bytes,setBytes]=useState<Uint8Array>(),[loading,setLoading]=useState(false),[readError,setReadError]=useState(''),[busy,setBusy]=useState(false),[error,setError]=useState(''),[progress,setProgress]=useState('');
 const [encoding,setEncoding]=useState<TextEncoding|''>(''),[delimiter,setDelimiter]=useState<','|'\t'>(()=>initialFiles?.length===1&&/\.tsv$/i.test(initialFiles[0].name)?'\t':','),[quoteMode,setQuoteMode]=useState<'csv'|'none'>('csv');
 const [failures,setFailures]=useState<{name:string;message:string}[]>(),[imported,setImported]=useState<string[]>([]);
 const file=files.length===1?files[0]:undefined,delimited=!!file&&officeDelimited(file.name);
 const problems=files.map(item=>({name:item.name,message:officeImportProblem(item.name,item.size)})).filter(item=>item.message);
 const batchProblem=files.length>MAX_BATCH?`每次最多导入 ${MAX_BATCH} 个文件`:files.length>1&&files.some(item=>officeDelimited(item.name))?'CSV/TSV 需要单独导入以选择编码和分隔符':'';
 const options:DelimitedImportOptions={...(encoding?{encoding}:{}),delimiter,quoteMode};
 useEffect(()=>{let active=true;setBytes(undefined);setReadError('');setLoading(false);if(!file||!delimited||officeImportProblem(file.name,file.size))return;setLoading(true);void file.arrayBuffer().then(data=>{if(active)setBytes(new Uint8Array(data));}).catch(()=>{if(active)setReadError('文件读取失败，请重新选择');}).finally(()=>{if(active)setLoading(false);});return()=>{active=false;};},[file,delimited]);
 const parsed=useMemo(()=>{if(!bytes)return undefined;try{return {value:prepareDelimitedImport(bytes,{...(encoding?{encoding}:{}),delimiter,quoteMode}),error:''};}catch(error){return {value:undefined,error:error instanceof Error?error.message:'无法解析 CSV/TSV'};}},[bytes,encoding,delimiter,quoteMode]);
 const ready=!!files.length&&!problems.length&&!batchProblem&&!loading&&!readError&&(!delimited||!!parsed?.value)&&!workspace.readOnly&&!failures;
 const choose=(next:File[])=>{setFiles(next);setBytes(undefined);setError('');setReadError('');setEncoding('');setDelimiter(next.length===1&&/\.tsv$/i.test(next[0].name)?'\t':',');setQuoteMode('csv');};
 const start=async()=>{if(!ready||busy)return;setBusy(true);setError('');const ids:string[]=[],failed:{name:string;message:string}[]=[];
  try{for(const [index,item]of files.entries()){if(files.length>1)setProgress(`正在导入 ${index+1}/${files.length}：${item.name}`);
   try{ids.push(await workspace.importOfficeFile(item,delimited?options:undefined));}catch(error){const message=error instanceof Error?error.message:'导入失败，请重试';if(files.length===1){setError(message);return;}failed.push({name:item.name,message});}}
   if(failed.length){setImported(ids);setFailures(failed);}else onFinish(ids);
  }finally{setBusy(false);setProgress('');}
 };
 const startRef=useRef(start);startRef.current=start;const autoStarted=useRef(false);
 // A drop is the import request; Word / Excel / PPT / PDF need no options, so start without another click.
 useEffect(()=>{if(autoStarted.current||!initialFiles?.length||!ready||delimited)return;autoStarted.current=true;void startRef.current();},[ready,delimited,initialFiles]);
 const {dragging,props:dropProps}=useFileDrop({disabled:busy||workspace.readOnly||!!failures,multiple:true,onFiles:choose,onError:setError});
 return createPortal(<Modal title="导入办公文件" className="delimited-export-dialog" onClose={()=>{if(!busy)onClose();}} drop={dropProps}><form onSubmit={event=>{event.preventDefault();void start();}}>
  <p>原件按原始字节保留，每个最多 20 MB。导入将创建新文件。</p>
  {!failures&&<FileDropHint dragging={dragging} disabled={busy||workspace.readOnly} hint="支持 Word、Excel、CSV/TSV、PowerPoint 和 PDF，可一次拖入多个文件"/>}
  <label className="form-field">办公文件<input aria-label="办公文件" type="file" multiple accept={ACCEPT} disabled={busy||workspace.readOnly||!!failures} onChange={event=>{choose([...(event.target.files??[])]);event.target.value='';}}/></label>
  {file&&<p className="office-backup-filename">{file.name} · {(file.size/1024).toFixed(1)} KB</p>}
  {files.length>1&&<ul className="import-preview" aria-label="待导入办公文件">{files.map((item,index)=><li key={index}><strong>{item.name}</strong><small>{(item.size/1024).toFixed(1)} KB</small></li>)}</ul>}
  {delimited&&<>
   <label className="form-field">文本编码<select aria-label="导入文本编码" value={encoding} disabled={busy} onChange={event=>{setEncoding(event.target.value as TextEncoding|'');setError('');}}><option value="">自动（BOM 优先，否则 UTF-8）</option><option value="utf-8">UTF-8</option><option value="utf-16le">UTF-16LE</option><option value="utf-16be">UTF-16BE</option></select></label>
   <label className="form-field">分隔符<select aria-label="导入分隔符" value={delimiter} disabled={busy} onChange={event=>{const value=event.target.value as ','|'\t';setDelimiter(value);if(value===',')setQuoteMode('csv');setError('');}}><option value=",">逗号</option><option value={'\t'}>制表符</option></select></label>
   <label className="form-field">引号模式<select aria-label="导入引号模式" value={quoteMode} disabled={busy} onChange={event=>{setQuoteMode(event.target.value as 'csv'|'none');setError('');}}><option value="csv">CSV 式双引号转义</option><option value="none" disabled={delimiter!== '\t'}>无引号转义（仅 TSV）</option></select></label>
   <p>首行保留为数据，不自动移除表头。前导零、长编号和公式样式文本均保留原文；所有导入单元格默认为文本，不计算 = 开头内容；导入后可为选中单元格选择公式类型。尚不支持日期等列类型转换。</p>
   {loading&&<p role="status">正在读取并解析文件…</p>}
   {parsed?.value&&<section aria-label="导入内容预览"><p role="status">{parsed.value.rows.length} 行 · {parsed.value.settings.encoding.toUpperCase()} · {parsed.value.settings.bom?'带 BOM':'无 BOM'} · 最多预览前 5 行、5 列，每格 120 字符</p>{parsed.value.rows.length?<div className="delimited-export-preview"><table><thead><tr>{Array.from({length:Math.min(5,Math.max(...parsed.value.rows.map(row=>row.length)))},(_,c)=><th scope="col" key={c}>第 {c+1} 列</th>)}</tr></thead><tbody>{parsed.value.rows.slice(0,5).map((row,r)=><tr key={r}>{row.slice(0,5).map((value,c)=><td key={c}>{value.slice(0,120)}{value.length>120?'…':''}</td>)}</tr>)}</tbody></table></div>:<p>文件为空，将保留原件并创建空工作表。</p>}</section>}
  </>}
  {progress&&<p role="status">{progress}</p>}
  {(problems.length>0||batchProblem||readError||parsed?.error||error)&&<p role="alert" className="form-error">{batchProblem||(problems.length?problems.map(item=>files.length>1?`${item.name}：${item.message}`:item.message).join('；'):'')||readError||parsed?.error||error}</p>}
  {failures&&<><p role="status">已导入 {imported.length} 个，未导入 {failures.length} 个。</p><ul className="import-preview" aria-label="办公文件导入结果">{failures.map((item,index)=><li key={index}><strong>{item.name} · 未导入</strong><small className="form-error">{item.message}</small></li>)}</ul></>}
  {workspace.readOnly&&<p>当前工作区只读，无法导入。</p>}
  <div className="dialog-actions">{failures?<button type="button" className="primary" onClick={()=>onFinish(imported)}>完成</button>:<><button type="button" onClick={onClose} disabled={busy}>取消</button><button className="primary" disabled={!ready||busy}>{busy?'正在导入…':files.length>1?`导入 ${files.length} 个文件`:'导入'}</button></>}</div>
 </form></Modal>,document.body);
}
