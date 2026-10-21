import {useEffect,useRef,useState} from 'react';
import {createPortal} from 'react-dom';
import {Modal} from './dialogs';
import {cloudJson,cloudRequest} from '../../../common/nbstore/src/cloud';
import {downloadOfficeBytes} from './office-download';
import type {LocalWorkspace} from './workspace';
type Job={id:string;user:string;userName:string|null;format:string;scope:string;sheet:string;status:'queued'|'running'|'done'|'failed';created:number;finished:number|null;snapshot:string;bytes:number;filename:string;mime:string;error:string|null};
const OPTIONS:Record<string,[string,string,string][]>={word:[['docx','paragraphs','DOCX 正文副本']],presentation:[['pptx','slides','PPTX 演示副本']],spreadsheet:[['xlsx','text-grid','XLSX 文本副本'],['xlsx','formula-grid','XLSX 公式副本'],['xlsx','formula-cache-grid','XLSX 公式结果副本'],['csv','escaped-text','CSV 文本'],['tsv','escaped-text','TSV 文本']]};
const STATUS={queued:'排队中',running:'正在生成',done:'已完成',failed:'失败'};
export const officeExportKinds=Object.keys(OPTIONS);
/** Persistent export history: direct downloads are recorded; background exports survive closing the page or restarting the service. */
export function OfficeExportHistoryDialog({workspace,id,onClose}:{workspace:LocalWorkspace;id:string;onClose:()=>void}){
 const resource=workspace.officeResources.get(id),options=OPTIONS[resource?.kind??'']??[],[jobs,setJobs]=useState<Job[]>([]),[choice,setChoice]=useState(0),[busy,setBusy]=useState(false),[error,setError]=useState(''),pending=useRef<{operationId:string;choice:number}>(),readOnly=workspace.documentReadOnly(id),query=new URLSearchParams({workspace:workspace.id,id}).toString();
 const load=async()=>{try{setJobs(await cloudJson<Job[]>('/office-exports?'+query));}catch(error){setError(error instanceof Error?error.message:'读取导出记录失败');}};
 useEffect(()=>{void load();const changed=(event:Event)=>{const detail=(event as CustomEvent<{workspace:string;id:string}>).detail;if(detail.workspace===workspace.id&&detail.id===id)void load();};window.addEventListener('office-exports-changed',changed);return()=>window.removeEventListener('office-exports-changed',changed);},[workspace,id]);
 // Local workspaces have no event stream, so poll while a job is still running.
 useEffect(()=>{if(!jobs.some(job=>job.status==='queued'||job.status==='running'))return;const timer=setInterval(()=>void load(),1500);return()=>clearInterval(timer);},[jobs]);
 const start=async()=>{const option=options[choice];if(!option||busy||readOnly)return;setBusy(true);setError('');if(pending.current?.choice!==choice)pending.current={operationId:crypto.randomUUID(),choice};try{await workspace.prepareOfficeBackup(id);await cloudJson('/office-exports?'+query,'POST',{operationId:pending.current.operationId,format:option[0],scope:option[1]});pending.current=undefined;await load();}catch(error){setError((error instanceof Error?error.message:'创建导出任务失败')+'；可重试，不会重复创建');}finally{setBusy(false);}};
 const download=async(job:Job)=>{setError('');try{const response=await cloudRequest('/office-exports?'+query+'&job='+encodeURIComponent(job.id));downloadOfficeBytes(job.filename,new Uint8Array(await response.arrayBuffer()),job.mime);}catch(error){setError(error instanceof Error?error.message:'下载失败');}};
 const remove=async(job:Job)=>{setError('');try{await cloudRequest('/office-exports?'+query+'&job='+encodeURIComponent(job.id),{method:'DELETE'});await load();}catch(error){setError(error instanceof Error?error.message:'删除失败');}};
 const label=(job:Job)=>options.find(([format,scope])=>format===job.format&&scope===job.scope)?.[2]??job.format.toUpperCase();
 return createPortal(<Modal title="导出记录" onClose={onClose}><p>每次导出都固定当时已保存的内容，记录保留生成的文件，可再次下载。后台导出可关闭页面，服务重启后会重新执行；每个文件保留最近 20 条，工作区共 200 MB。</p>
  {error&&<p role="alert">{error}</p>}
  {options.length>0&&<div className="dialog-actions"><label>导出类型<select aria-label="后台导出类型" value={choice} disabled={busy} onChange={event=>setChoice(Number(event.target.value))}>{options.map((option,i)=><option key={option[2]} value={i}>{option[2]}</option>)}</select></label><button className="primary" disabled={busy} onClick={()=>void start()}>{busy?'正在创建…':'后台导出'}</button></div>}
  {!jobs.length?<p>暂无导出记录。</p>:<ol className="office-export-history">{jobs.map(job=><li key={job.id} data-status={job.status}><div><strong>{label(job)}</strong> · <span role="status">{STATUS[job.status]}</span>{job.status==='done'&&<> · {(job.bytes/1024).toFixed(1)} KB</>}</div><small>{new Date(job.created).toLocaleString()} · {job.userName??'本人'} · 快照 {job.snapshot.slice(0,8)}</small>{job.error&&<p className="form-error">{job.error}</p>}<div className="office-comment-actions"><button disabled={job.status!=='done'} onClick={()=>void download(job)}>下载</button><button onClick={()=>void remove(job)}>删除记录</button></div></li>)}</ol>}
  <div className="dialog-actions"><button onClick={onClose}>关闭</button></div>
 </Modal>,document.body);
}
