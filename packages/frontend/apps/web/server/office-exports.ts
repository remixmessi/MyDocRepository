import {createHash,randomUUID} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import type {OfficeKind} from '../../../../common/nbstore/src/office';
import {packOfficeMedia,readResourceMedia} from './office-media';
import type {OfficeParserPool} from './office-parser-process';
export const OFFICE_EXPORT_LIMITS={perResource:20,active:5,workspaceBytes:200*1024*1024,outputBytes:50*1024*1024};
export class OfficeExportError extends Error{constructor(message:string,readonly status=400){super(message);}}
export type OfficeExportPlan={kind:OfficeKind;format:string;scope:string;sheet:string;media:boolean;suffix:string;mime:string;calculated:boolean};
const MIME={docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation',csv:'text/csv; charset=utf-8',tsv:'text/tab-separated-values; charset=utf-8'} as const;
/** The one place that maps a requested export to a resource kind, file name and parser task. */
export function officeExportPlan(format:string|null,scope:string|null,sheet='first'):OfficeExportPlan{
 const s=scope??'',delimited=(format==='csv'||format==='tsv')&&['escaped-text','raw-text','escaped-values','raw-values'].includes(s);
 if(!/^[a-zA-Z0-9-]{1,128}$/.test(sheet))throw new OfficeExportError('工作表编号无效');
 if(format==='pptx'&&s==='slides')return {kind:'presentation',format,scope:s,sheet:'',media:true,suffix:'-演示副本.pptx',mime:MIME.pptx,calculated:false};
 if(format==='docx'&&s==='paragraphs')return {kind:'word',format,scope:s,sheet:'',media:true,suffix:'-正文副本.docx',mime:MIME.docx,calculated:false};
 if(format==='xlsx'&&['text-grid','formula-grid','formula-cache-grid'].includes(s))return {kind:'spreadsheet',format,scope:s,sheet:'',media:false,suffix:s==='text-grid'?'-文本副本.xlsx':s==='formula-grid'?'-公式副本.xlsx':'-公式结果副本.xlsx',mime:MIME.xlsx,calculated:s==='formula-cache-grid'};
 if(delimited)return {kind:'spreadsheet',format:format!,scope:s,sheet,media:false,suffix:(s.endsWith('-values')?'-计算结果.':'-文本副本.')+format,mime:MIME[format as 'csv'|'tsv'],calculated:s.endsWith('-values')};
 throw new OfficeExportError('请明确确认办公文件兼容副本的导出范围');
}
export function runOfficeExport(parser:OfficeParserPool,plan:OfficeExportPlan,input:Uint8Array,signal?:AbortSignal){
 if(plan.format==='pptx')return parser.exportPptx(input,signal);if(plan.format==='docx')return parser.exportDocx(input,signal);
 if(plan.format==='xlsx')return plan.scope==='text-grid'?parser.exportXlsxText(input,signal):parser.exportXlsxFormula(input,signal,plan.scope==='formula-cache-grid');
 return parser.exportDelimited(input,plan.format as 'csv'|'tsv',plan.scope.startsWith('escaped-')?'escaped':'raw',signal,plan.scope.endsWith('-values')?'values':'text',plan.sheet||'first');
}
/** Read the current saved state once; the export only ever sees this pinned snapshot. */
export function captureOfficeExport(database:SQLiteProjectDatabase,workspace:string,id:string,plan:OfficeExportPlan){
 const row=database.db.prepare('SELECT r.name,r.kind,d.data FROM cloud_office_resources r JOIN documents d ON d.scope=? AND d.id=r.id WHERE r.workspace=? AND r.id=? AND r.deleted=0').get('cloud:'+workspace,workspace,id) as {name:string;kind:string;data:Buffer}|undefined;
 if(!row||row.kind!==plan.kind)throw new OfficeExportError('办公文件不存在或类型不匹配',404);
 const snapshot=Buffer.from(row.data),input=plan.media?packOfficeMedia(snapshot,readResourceMedia(database,workspace,id,row.kind as OfficeKind,snapshot)):snapshot;
 return {snapshot,input,hash:createHash('sha256').update(snapshot).digest('hex'),filename:row.name.replace(/\.[^.]+$/,'')+plan.suffix};
}
export type OfficeExportJob={id:string;user:string;format:string;scope:string;sheet:string;status:'queued'|'running'|'done'|'failed';created:number;finished:number|null;snapshot:string;sha256:string|null;bytes:number;filename:string;mime:string;error:string|null};
const columns='id,user,format,scope,sheet,status,created,finished,snapshot,sha256,bytes,filename,mime,error';
function prune(database:SQLiteProjectDatabase,workspace:string,resource:string){
 const finished=database.db.prepare("SELECT id FROM cloud_office_exports WHERE workspace=? AND resource=? AND status IN ('done','failed') ORDER BY created DESC,id DESC").all(workspace,resource) as {id:string}[];
 const active=(database.db.prepare("SELECT count(*) n FROM cloud_office_exports WHERE workspace=? AND resource=? AND status IN ('queued','running')").get(workspace,resource) as {n:number}).n;
 for(const row of finished.slice(Math.max(0,OFFICE_EXPORT_LIMITS.perResource-active)))database.db.prepare('DELETE FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?').run(workspace,resource,row.id);
 // Oldest finished outputs go first when the workspace exceeds its export storage budget.
 let total=(database.db.prepare('SELECT coalesce(sum(bytes),0) n FROM cloud_office_exports WHERE workspace=?').get(workspace) as {n:number}).n;
 for(const row of database.db.prepare("SELECT resource,id,bytes FROM cloud_office_exports WHERE workspace=? AND status='done' ORDER BY created,id").all(workspace) as {resource:string;id:string;bytes:number}[]){if(total<=OFFICE_EXPORT_LIMITS.workspaceBytes)break;database.db.prepare('DELETE FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?').run(workspace,row.resource,row.id);total-=row.bytes;}
}
/** Record a finished direct download so it appears in the file's export history. */
export function recordOfficeExport(database:SQLiteProjectDatabase,workspace:string,resource:string,user:string,plan:OfficeExportPlan,captured:{hash:string;filename:string},output:Uint8Array){
 if(output.length>OFFICE_EXPORT_LIMITS.outputBytes)return;database.db.transaction(()=>{if(!database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,resource))return;const now=Date.now();
  database.db.prepare("INSERT INTO cloud_office_exports(workspace,resource,id,user,operation,format,scope,sheet,status,created,finished,snapshot,input,output,sha256,bytes,filename,mime,error) VALUES (?,?,?,?,?,?,?,?,'done',?,?,?,NULL,?,?,?,?,?,NULL)").run(workspace,resource,randomUUID(),user,'direct-'+randomUUID(),plan.format,plan.scope,plan.sheet,now,now,captured.hash,Buffer.from(output),createHash('sha256').update(output).digest('hex'),output.length,captured.filename,plan.mime);prune(database,workspace,resource);}).immediate();
}
/** Queue a background export from a pinned snapshot; the same operation id returns the same job. */
export function createOfficeExportJob(database:SQLiteProjectDatabase,workspace:string,resource:string,user:string,operation:string,plan:OfficeExportPlan){
 if(!/^[a-zA-Z0-9-]{16,80}$/.test(operation))throw new OfficeExportError('导出任务编号无效');
 return database.db.transaction(()=>{
  const existing=database.db.prepare(`SELECT ${columns} FROM cloud_office_exports WHERE workspace=? AND resource=? AND user=? AND operation=?`).get(workspace,resource,user,operation) as OfficeExportJob|undefined;
  if(existing){if(existing.format!==plan.format||existing.scope!==plan.scope||existing.sheet!==plan.sheet)throw new OfficeExportError('导出任务编号已用于其他范围',409);return {job:existing,repeated:true};}
  if((database.db.prepare("SELECT count(*) n FROM cloud_office_exports WHERE workspace=? AND resource=? AND status IN ('queued','running')").get(workspace,resource) as {n:number}).n>=OFFICE_EXPORT_LIMITS.active)throw new OfficeExportError('此文件已有 5 个导出任务在进行，请稍后再试',429);
  const captured=captureOfficeExport(database,workspace,resource,plan),id=randomUUID(),now=Date.now();
  database.db.prepare("INSERT INTO cloud_office_exports(workspace,resource,id,user,operation,format,scope,sheet,status,created,finished,snapshot,input,output,sha256,bytes,filename,mime,error) VALUES (?,?,?,?,?,?,?,?,'queued',?,NULL,?,?,NULL,NULL,0,?,?,NULL)").run(workspace,resource,id,user,operation,plan.format,plan.scope,plan.sheet,now,captured.hash,Buffer.from(captured.input),captured.filename,plan.mime);prune(database,workspace,resource);
  return {job:database.db.prepare(`SELECT ${columns} FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?`).get(workspace,resource,id) as OfficeExportJob,repeated:false};
 }).immediate();
}
export function listOfficeExports(database:SQLiteProjectDatabase,workspace:string,resource:string){return database.db.prepare(`SELECT ${columns.split(',').map(c=>'e.'+c).join(',')},u.name AS userName FROM cloud_office_exports e LEFT JOIN cloud_users u ON u.id=e.user WHERE e.workspace=? AND e.resource=? ORDER BY e.created DESC,e.id DESC`).all(workspace,resource) as (OfficeExportJob&{userName:string|null})[];}
export function readOfficeExport(database:SQLiteProjectDatabase,workspace:string,resource:string,id:string){
 const row=database.db.prepare('SELECT status,output,sha256,filename,mime,snapshot FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?').get(workspace,resource,id) as {status:string;output:Buffer|null;sha256:string|null;filename:string;mime:string;snapshot:string}|undefined;
 if(!row)throw new OfficeExportError('导出记录不存在或已清理',404);if(row.status!=='done'||!row.output)throw new OfficeExportError(row.status==='failed'?'导出任务失败，请重新导出':'导出任务尚未完成',409);
 if(createHash('sha256').update(row.output).digest('hex')!==row.sha256)throw new OfficeExportError('导出产物校验失败，请重新导出',500);return row as typeof row&{output:Buffer;sha256:string};
}
export function deleteOfficeExport(database:SQLiteProjectDatabase,workspace:string,resource:string,id:string,user:string,admin:boolean){
 const row=database.db.prepare('SELECT user FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?').get(workspace,resource,id) as {user:string}|undefined;if(!row)throw new OfficeExportError('导出记录不存在或已清理',404);if(row.user!==user&&!admin)throw new OfficeExportError('只有发起人或管理员可删除导出记录',403);
 database.db.prepare('DELETE FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?').run(workspace,resource,id);
}
/** Runs queued jobs one at a time. Jobs left running by a crash or restore are re-queued on start. */
export class OfficeExportQueue{
 private active=false;private stopped=false;private controller?:AbortController;
 constructor(private readonly database:SQLiteProjectDatabase,private readonly parser:OfficeParserPool,private readonly notify:(workspace:string,resource:string)=>void=()=>{}){
  database.db.prepare("UPDATE cloud_office_exports SET status='queued' WHERE status='running'").run();
 }
 kick(){if(!this.active&&!this.stopped)void this.drain();}
 close(){this.stopped=true;this.controller?.abort();}
 private async drain(){
  this.active=true;try{while(!this.stopped){
   const job=this.database.db.prepare("SELECT workspace,resource,id,format,scope,sheet,input FROM cloud_office_exports WHERE status='queued' ORDER BY created,id LIMIT 1").get() as {workspace:string;resource:string;id:string;format:string;scope:string;sheet:string;input:Buffer|null}|undefined;if(!job)break;
   const where=[job.workspace,job.resource,job.id];this.database.db.prepare("UPDATE cloud_office_exports SET status='running' WHERE workspace=? AND resource=? AND id=?").run(...where);
   let output:Uint8Array|undefined,error:string|undefined;this.controller=new AbortController();
   try{if(!job.input)throw new Error('导出输入缺失，请重新导出');output=await runOfficeExport(this.parser,officeExportPlan(job.format,job.scope,job.sheet||'first'),job.input,this.controller.signal);if(output.length>OFFICE_EXPORT_LIMITS.outputBytes)throw new Error('导出产物超过 50 MB');}catch(failure){
    // A busy parser pool is not a job failure: put it back and try again shortly, leaving room for interactive requests.
    if((failure as {status?:number}).status===503&&!this.stopped){this.database.db.prepare("UPDATE cloud_office_exports SET status='queued' WHERE workspace=? AND resource=? AND id=? AND status='running'").run(...where);this.active=false;this.controller=undefined;setTimeout(()=>this.kick(),1000).unref?.();return;}
    error=failure instanceof Error?failure.message.slice(0,500):'导出失败';}
   if(this.stopped&&!output)break;
   this.database.db.transaction(()=>{
    if(!this.database.db.prepare('SELECT 1 FROM cloud_office_exports WHERE workspace=? AND resource=? AND id=?').get(...where))return;
    const live=!!this.database.db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(job.workspace,job.resource);
    if(output&&live)this.database.db.prepare("UPDATE cloud_office_exports SET status='done',finished=?,input=NULL,output=?,sha256=?,bytes=?,error=NULL WHERE workspace=? AND resource=? AND id=?").run(Date.now(),Buffer.from(output),createHash('sha256').update(output).digest('hex'),output.length,...where);
    else this.database.db.prepare("UPDATE cloud_office_exports SET status='failed',finished=?,input=NULL,error=? WHERE workspace=? AND resource=? AND id=?").run(Date.now(),live?error??'导出失败':'文件已删除',...where);
    prune(this.database,job.workspace,job.resource);
   }).immediate();this.notify(job.workspace,job.resource);
  }}finally{this.active=false;this.controller=undefined;}
 }
}
