import {useEffect,useMemo,useRef,useState} from 'react';
import {Upload} from 'lucide-react';
import {Modal} from './dialogs';
import type {LocalWorkspace} from './workspace';
import {inspectOfficeText,MAX_OFFICE_FILE_BYTES,type TextEncoding,type TextLineEnding} from '../../../common/nbstore/src/office';

export function TextImportDialog({workspace,onClose,onFinish}:{workspace:LocalWorkspace;onClose:()=>void;onFinish:(id:string)=>void}){
  const picker=useRef<HTMLInputElement>(null);
  const [file,setFile]=useState<File>(),[loaded,setLoaded]=useState<{file:File;bytes:Uint8Array}>();
  const [encoding,setEncoding]=useState<TextEncoding|'auto'>('auto');
  const [lineEnding,setLineEnding]=useState<TextLineEnding|''>('');
  const [busy,setBusy]=useState(false),[error,setError]=useState('');
  useEffect(()=>{
    if(!file)return;let active=true;setLoaded(undefined);setError('');
    if(file.size>MAX_OFFICE_FILE_BYTES){setError('原件超过 20 MB');return;}
    if(!/\.txt$/i.test(file.name)){setError('请选择 TXT 文件');return;}
    void file.arrayBuffer().then(buffer=>{if(active)setLoaded({file,bytes:new Uint8Array(buffer)});}).catch(()=>{if(active)setError('无法读取文件，请重新选择');});
    return()=>{active=false;};
  },[file]);
  const preview=useMemo(()=>{
    if(!loaded||loaded.file!==file)return undefined;
    try{return {value:inspectOfficeText(loaded.bytes,encoding==='auto'?undefined:encoding)};}
    catch(error){return {error:error instanceof TypeError?'无法按所选编码解码，请选择正确的 UTF 编码。':error instanceof Error?error.message:'无法解码文件'};}
  },[file,loaded,encoding]);
  const readOnly=workspace.readOnly;
  return <Modal title="导入 TXT" wide onClose={()=>{if(!busy)onClose();}}>
    <form onSubmit={event=>{event.preventDefault();if(!file||!preview?.value||busy||readOnly)return;setBusy(true);setError('');void workspace.importOfficeText(file,encoding==='auto'?undefined:encoding,preview.value.mixed?lineEnding||undefined:undefined).then(onFinish).catch(error=>setError(error instanceof Error?error.message:'导入失败，请重试')).finally(()=>setBusy(false));}}>
      <div className="import-file-actions"><button type="button" disabled={busy||readOnly} onClick={()=>picker.current?.click()}><Upload size={16}/>选择 TXT 文件</button><span className="text-import-filename" title={file?.name}>{file?.name??'尚未选择文件'}</span><input ref={picker} type="file" hidden accept=".txt,text/plain" aria-label="导入 TXT 文件" disabled={busy||readOnly} onChange={event=>{const file=event.target.files?.[0];event.target.value='';if(file){setEncoding('auto');setLineEnding('');setFile(file);}}}/></div>
      <div className="text-format-controls"><label>文件编码<select aria-label="导入文本编码" value={encoding} disabled={busy} onChange={event=>setEncoding(event.target.value as TextEncoding|'auto')}><option value="auto">自动（BOM / UTF-8）</option><option value="utf-8">UTF-8</option><option value="utf-16le">UTF-16LE</option><option value="utf-16be">UTF-16BE</option></select></label>
      {preview?.value&&<span role="status">{preview.value.encoding.toUpperCase()}{preview.value.bom?' · BOM':''} · {preview.value.mixed?'混合换行':preview.value.lineEnding.toUpperCase()}</span>}
      {preview?.value?.mixed&&<label>换行规范化<select required aria-label="混合换行处理" value={lineEnding} disabled={busy} onChange={event=>setLineEnding(event.target.value as TextLineEnding|typeof lineEnding)}><option value="">请选择</option><option value="lf">LF</option><option value="crlf">CRLF</option><option value="cr">CR</option></select></label>}</div>
      {file&&!loaded&&!error&&<p role="status">正在读取文件…</p>}
      {(error||preview?.error)&&<p role="alert" className="form-error">{error||preview?.error}</p>}
      {preview?.value&&<pre className="text-import-preview" aria-label="文本导入预览">{preview.value.text.slice(0,4000)||'（空文件）'}</pre>}
      <div className="dialog-actions"><button type="button" disabled={busy} onClick={onClose}>取消</button><button className="primary" disabled={busy||readOnly||!preview?.value||preview.value.mixed&&!lineEnding}>{busy?'正在导入…':'导入并保留原件'}</button></div>
    </form>
  </Modal>;
}
