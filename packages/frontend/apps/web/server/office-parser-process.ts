import {MAX_WORKSPACE_ARCHIVE_BYTES} from '../../../../common/nbstore/src/workspace-archive';
import {wrapPdfFont} from './office-pdf-font';
import {loadPdfFont} from './office-pdf-font-file';
import {validDelimitedImportOptions,type DelimitedImportOptions} from '../../../../common/nbstore/src/office-delimited-import';
import {spawn} from 'node:child_process';
import {MAX_OFFICE_FILE_BYTES} from '../../../../common/nbstore/src/office';
import type {OfficeParserFormat} from './office-parser-model';

export type OfficeJobOptions=DelimitedImportOptions|{sheetId:string}|{targetWorkspace?:string};
export type OfficeWorkerTask=OfficeParserFormat|'pdf-form-spec'|'pdf-native-annotations'|'export-pdf-flattened'|'export-pdf-annotations'|'personal-workspace-archive'|'workspace-archive'|'import-delimited'|'export-docx'|'export-pptx'|'export-xlsx-text'|'export-xlsx-formula'|'export-xlsx-formula-cache'|`export-${'csv'|'tsv'}-${'escaped'|'raw'}${''|'-values'}`|`validate-${'text'|'word'|'spreadsheet'|'pdf'|'presentation'}`;

declare const __MYDOC_OFFICE_PARSER_SOURCE__:string;
// sourceBytes bounds the trusted parser bundle sent on stdin; @pdf-lib/fontkit (PDF Unicode font) adds about 2 MB.
export const OFFICE_PARSER_LIMITS={timeoutMs:15000,heapMb:256,concurrency:2,outputBytes:16*1024*1024,sourceBytes:6*1024*1024};
export class OfficeParserError extends Error {constructor(message:string,readonly status=400){super(message);}}
async function parserSource(){
 if(typeof __MYDOC_OFFICE_PARSER_SOURCE__==='string'){
  return __MYDOC_OFFICE_PARSER_SOURCE__;
 }else{
  // Rebuild in development so hot reload cannot retain an older parser implementation.
  const {buildDevelopmentParser}=await import('./office-parser-development');
  return buildDevelopmentParser();
 }
}
// The bootstrap is small enough for OS argv limits. The trusted bundle and binary
// input travel on stdin with a length prefix, never through shell interpolation.
const BOOTSTRAP=`
let chunks=[],size=0;
process.stdin.on('data',chunk=>{size+=chunk.length;if(size>${MAX_WORKSPACE_ARCHIVE_BYTES+OFFICE_PARSER_LIMITS.sourceBytes+4})process.exit(1);chunks.push(chunk);});
process.stdin.on('end',()=>{
 const packet=Buffer.concat(chunks);chunks=[];
 if(packet.length<4)process.exit(1);
 const length=packet.readUInt32LE(0);
 if(length>${OFFICE_PARSER_LIMITS.sourceBytes}||length+4>packet.length||packet.length-length-4>((process.argv[1].endsWith('workspace-archive')||process.argv[1].startsWith('export-pdf-'))?${MAX_WORKSPACE_ARCHIVE_BYTES}:${MAX_OFFICE_FILE_BYTES}))process.exit(1);
 globalThis.__MYDOC_OFFICE_JOB__={format:process.argv[1],bytes:packet.subarray(4+length),options:JSON.parse(process.argv[2]??'null')};
 const mod={exports:{}};
 new Function('require','module','exports',packet.subarray(4,4+length).toString())(require,mod,mod.exports);
});`;

export function runOfficeParserProcess(source:string,format:OfficeWorkerTask,bytes:Uint8Array,options:{signal?:AbortSignal;timeoutMs?:number;outputBytes?:number;delimited?:OfficeJobOptions}={}){
 if(!['pdf-form-spec','pdf-native-annotations','export-pdf-flattened','export-pdf-annotations','personal-workspace-archive','workspace-archive','docx','xlsx','pptx','csv','tsv','import-delimited','export-docx','export-pptx','export-xlsx-text','export-xlsx-formula','export-xlsx-formula-cache','export-csv-escaped','export-csv-raw','export-tsv-escaped','export-tsv-raw','export-csv-escaped-values','export-csv-raw-values','export-tsv-escaped-values','export-tsv-raw-values','validate-text','validate-word','validate-spreadsheet','validate-pdf','validate-presentation'].includes(format))return Promise.reject(new OfficeParserError('办公解析格式无效'));
 if(format==='import-delimited'&&!validDelimitedImportOptions(options.delimited))return Promise.reject(new OfficeParserError('CSV/TSV 导入设置无效'));
 if(options.signal?.aborted)return Promise.reject(new OfficeParserError('文件解析已取消',499));
 if(bytes.length>((format.endsWith('workspace-archive')||format.startsWith('export-pdf-'))?MAX_WORKSPACE_ARCHIVE_BYTES:MAX_OFFICE_FILE_BYTES))return Promise.reject(new OfficeParserError((format.endsWith('workspace-archive')||format.startsWith('export-pdf-'))?'完整备份超过 50 MB':'原件超过 20 MB'));
 const code=Buffer.from(source);if(code.length>OFFICE_PARSER_LIMITS.sourceBytes)return Promise.reject(new OfficeParserError('解析程序超过大小限制',503));
 return new Promise<Uint8Array>((resolve,reject)=>{
  // Empty fs permissions; no child/worker/WASI permissions and no inherited secrets or NODE_OPTIONS.
  const child=spawn(process.execPath,[process.versions.electron?'--permission':'--experimental-permission',`--max-old-space-size=${((format.endsWith('workspace-archive')||format.startsWith('export-pdf-'))?512:OFFICE_PARSER_LIMITS.heapMb)}`,'-e',BOOTSTRAP,format,JSON.stringify(options.delimited??null)],{stdio:['pipe','pipe','pipe'],env:{NODE_ENV:'production',TZ:'UTC',...(process.versions.electron?{ELECTRON_RUN_AS_NODE:'1'}:{})},windowsHide:true});
  const chunks:Buffer[]=[];let size=0,stderrSize=0,failure:OfficeParserError|undefined;
  const fail=(message:string,status:number)=>{failure??=new OfficeParserError(message,status);child.kill('SIGKILL');};
  const abort=()=>fail('文件解析已取消',499);
  options.signal?.addEventListener('abort',abort,{once:true});
  if(options.signal?.aborted)abort();
  const timer=setTimeout(()=>fail('文件解析超时，请简化文件后重试',422),options.timeoutMs??OFFICE_PARSER_LIMITS.timeoutMs);
  child.stdout.on('data',(chunk:Buffer)=>{size+=chunk.length;if(size>(options.outputBytes??((format.endsWith('workspace-archive')||format.startsWith('export-pdf-'))?MAX_WORKSPACE_ARCHIVE_BYTES:OFFICE_PARSER_LIMITS.outputBytes))+1)fail('解析结果超过大小限制',413);else chunks.push(Buffer.from(chunk));});
  child.stderr.on('data',(chunk:Buffer)=>{stderrSize+=chunk.length;if(stderrSize>65536)fail('文件解析进程异常退出',503);});
  child.on('error',()=>{failure??=new OfficeParserError('无法启动办公解析进程',503);});
  child.stdin.on('error',()=>{failure??=new OfficeParserError('解析输入传输失败',503);});
  child.on('close',(exitCode)=>{
   clearTimeout(timer);options.signal?.removeEventListener('abort',abort);
   if(failure){reject(failure);return;}
   const output=Buffer.concat(chunks);
   if(exitCode!==0||!output.length){reject(new OfficeParserError('文件解析进程异常退出，原件尚未写入',503));return;}
   if(output[0]===1){reject(new OfficeParserError(output.subarray(1).toString('utf8').slice(0,500)||'办公文件解析失败'));return;}
   if(output[0]!==0||output.length<3){reject(new OfficeParserError('解析结果格式无效',503));return;}
   resolve(output.subarray(1));
  });
  const header=Buffer.alloc(4);header.writeUInt32LE(code.length);
  child.stdin.write(header);child.stdin.write(code);child.stdin.end(bytes);
 });
}

export class OfficeParserPool {
 private jobs=new Set<AbortController>();private closed=false;
 constructor(private options:{source?:()=>Promise<string>;timeoutMs?:number}={}){}
 pdfFormSpecs(bytes:Uint8Array,signal?:AbortSignal){return this.execute('pdf-form-spec',bytes,signal);}
 pdfNativeAnnotations(bytes:Uint8Array,signal?:AbortSignal){return this.execute('pdf-native-annotations',bytes,signal);}
 preparePersonalArchive(bytes:Uint8Array,targetWorkspace:string,signal?:AbortSignal){return this.execute('personal-workspace-archive',bytes,signal,{targetWorkspace});}
 prepareWorkspaceArchive(bytes:Uint8Array,targetWorkspace?:string,signal?:AbortSignal){return this.execute('workspace-archive',bytes,signal,{targetWorkspace});}
 // The parser process cannot read files, so the bundled Unicode font travels with the job.
 exportAnnotatedPdf(bytes:Uint8Array,signal?:AbortSignal,flatten=false){return this.execute(flatten?'export-pdf-flattened':'export-pdf-annotations',wrapPdfFont(bytes,loadPdfFont()),signal);}
 get activeCount(){return this.jobs.size;}
 parse(format:OfficeParserFormat,bytes:Uint8Array,signal?:AbortSignal){return this.execute(format,bytes,signal);}
 exportPptx(bytes:Uint8Array,signal?:AbortSignal){return this.execute('export-pptx',bytes,signal);}
 exportDocx(bytes:Uint8Array,signal?:AbortSignal){return this.execute('export-docx',bytes,signal);}
 exportXlsxText(bytes:Uint8Array,signal?:AbortSignal){return this.execute('export-xlsx-text',bytes,signal);}
 exportXlsxFormula(bytes:Uint8Array,signal?:AbortSignal,cache=false){return this.execute(cache?'export-xlsx-formula-cache':'export-xlsx-formula',bytes,signal);}
 exportDelimited(bytes:Uint8Array,format:import('../../../../common/nbstore/src/office-delimited-export').DelimitedFormat,mode:import('../../../../common/nbstore/src/office-delimited-export').DelimitedTextMode,signal?:AbortSignal,content:'text'|'values'='text',sheetId='first'){return this.execute(`export-${format}-${mode}${content==='values'?'-values':''}`,bytes,signal,{sheetId});}
 parseDelimited(bytes:Uint8Array,options:DelimitedImportOptions,signal?:AbortSignal){return this.execute('import-delimited',bytes,signal,options);}
 validateBackup(kind:import('../../../../common/nbstore/src/office').OfficeKind,bytes:Uint8Array,signal?:AbortSignal){return this.execute(`validate-${kind}`,bytes,signal);}
 private async execute(format:OfficeWorkerTask,bytes:Uint8Array,signal?:AbortSignal,delimited?:OfficeJobOptions){
  if(this.closed)throw new OfficeParserError('办公解析服务已关闭',503);
  if(!['pdf-form-spec','pdf-native-annotations','export-pdf-flattened','export-pdf-annotations','personal-workspace-archive','workspace-archive','docx','xlsx','pptx','csv','tsv','import-delimited','export-docx','export-pptx','export-xlsx-text','export-xlsx-formula','export-xlsx-formula-cache','export-csv-escaped','export-csv-raw','export-tsv-escaped','export-tsv-raw','export-csv-escaped-values','export-csv-raw-values','export-tsv-escaped-values','export-tsv-raw-values','validate-text','validate-word','validate-spreadsheet','validate-pdf','validate-presentation'].includes(format))throw new OfficeParserError('办公解析格式无效');
  if(bytes.length>((format.endsWith('workspace-archive')||format.startsWith('export-pdf-'))?MAX_WORKSPACE_ARCHIVE_BYTES:MAX_OFFICE_FILE_BYTES))throw new OfficeParserError((format.endsWith('workspace-archive')||format.startsWith('export-pdf-'))?'完整备份超过 50 MB':'原件超过 20 MB');
  if(signal?.aborted)throw new OfficeParserError('文件解析已取消',499);
  if(this.jobs.size>=OFFICE_PARSER_LIMITS.concurrency)throw new OfficeParserError('正在处理其他办公文件，请稍后重试',503);
  const controller=new AbortController(),abort=()=>controller.abort();this.jobs.add(controller);signal?.addEventListener('abort',abort,{once:true});
  const start=Date.now(),timeoutMs=this.options.timeoutMs??OFFICE_PARSER_LIMITS.timeoutMs;
  try{
   let timer:ReturnType<typeof setTimeout>|undefined,rejectPreparation:(reason:unknown)=>void=()=>{};
   const onAbort=()=>rejectPreparation(new OfficeParserError('文件解析已取消',499));
   const cancelled=new Promise<never>((_,reject)=>{rejectPreparation=reject;timer=setTimeout(()=>reject(new OfficeParserError('文件解析超时，请简化文件后重试',422)),timeoutMs);});
   void cancelled.catch(()=>{});controller.signal.addEventListener('abort',onAbort,{once:true});
   let source:string;
   try{
    if(controller.signal.aborted)throw new OfficeParserError('文件解析已取消',499);
    source=await Promise.race([(this.options.source??parserSource)(),cancelled]);
   }catch(error){if(error instanceof OfficeParserError)throw error;throw new OfficeParserError('无法准备办公解析进程',503);}
   finally{clearTimeout(timer);controller.signal.removeEventListener('abort',onAbort);}
   if(controller.signal.aborted)throw new OfficeParserError('文件解析已取消',499);
   const update=await runOfficeParserProcess(source,format,bytes,{signal:controller.signal,timeoutMs:Math.max(1,timeoutMs-(Date.now()-start)),delimited});
   if(controller.signal.aborted)throw new OfficeParserError('文件解析已取消',499);return update;
  }finally{signal?.removeEventListener('abort',abort);this.jobs.delete(controller);}
 }
 close(){this.closed=true;for(const job of this.jobs)job.abort();}
}
