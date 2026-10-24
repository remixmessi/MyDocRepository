import {preparePersonalWorkspaceArchive} from './personal-workspace-archive-model';
import {inspectPdfForms} from './office-pdf-forms';
import {readNativeAnnotations} from './office-pdf-native';
import {PDFDocument} from 'pdf-lib';
import {exportAnnotatedPdf} from './office-pdf-export';
import {unwrapPdfFont} from './office-pdf-font';
import {prepareWorkspaceArchive} from './workspace-archive-model';
import {exportPptxPresentation} from './office-presentation-export';
import {encodeDelimitedOriginal,validateSpreadsheetUpdate} from './office-spreadsheet';
import type {DelimitedImportOptions} from '../../../../common/nbstore/src/office-delimited-import';
import {exportDelimitedText} from './office-delimited-export';
import type {DelimitedFormat,DelimitedTextMode} from '../../../../common/nbstore/src/office-delimited-export';
import {exportXlsxText,exportXlsxFormula} from './office-spreadsheet-export';
import {validateOfficeBackupModel} from './office-backup-model';
import type {OfficeKind} from '../../../../common/nbstore/src/office';
import {exportDocxWord} from './office-word-export';
import type {OfficeWorkerTask} from './office-parser-process';
import {parseOfficeModel} from './office-parser-model';
// Only the parent-created bootstrap supplies this value. File data is never executed.
declare const __MYDOC_OFFICE_JOB__:{format:OfficeWorkerTask;bytes:Uint8Array;options:import('./office-parser-process').OfficeJobOptions};
void(async()=>{try{
 const {format,bytes,options}=__MYDOC_OFFICE_JOB__;
 const importDelimited=()=>{const parsed=encodeDelimitedOriginal(bytes,options as DelimitedImportOptions);try{validateSpreadsheetUpdate(null,parsed.update);if(parsed.update.length>16*1024*1024)throw new Error('协作模型超过 16 MB');return parsed.update;}finally{parsed.doc.destroy();}};
 const update=format==='pdf-form-spec'?await inspectPdfForms(bytes):format==='pdf-native-annotations'?new TextEncoder().encode(JSON.stringify(readNativeAnnotations(await PDFDocument.load(bytes,{updateMetadata:false,throwOnInvalidObject:true})))):format.startsWith('export-pdf-')?await (({packet,font})=>exportAnnotatedPdf(packet,format==='export-pdf-flattened',font,(options as {pages?:number[]}|undefined)?.pages))(unwrapPdfFont(bytes)):format==='personal-workspace-archive'?preparePersonalWorkspaceArchive(bytes,(options as {targetWorkspace:string}).targetWorkspace):format==='workspace-archive'?prepareWorkspaceArchive(bytes,(options as {targetWorkspace?:string}|undefined)?.targetWorkspace):format==='import-delimited'?importDelimited():/^export-(csv|tsv)-(escaped|raw)(-values)?$/.test(format)?exportDelimitedText(bytes,format.split('-')[1] as DelimitedFormat,format.split('-')[2] as DelimitedTextMode,format.endsWith('-values')?'values':'text',(options as {sheetId?:string}|undefined)?.sheetId):format==='export-xlsx-text'?exportXlsxText(bytes):format==='export-xlsx-formula'||format==='export-xlsx-formula-cache'?exportXlsxFormula(bytes,format==='export-xlsx-formula-cache'):format==='export-pptx'?exportPptxPresentation(bytes):format==='export-docx'?exportDocxWord(bytes):format.startsWith('validate-')?validateOfficeBackupModel(format.slice(9) as OfficeKind,bytes):parseOfficeModel(format as import('./office-parser-model').OfficeParserFormat,bytes);
 process.stdout.end(Buffer.concat([Buffer.from([0]),Buffer.from(update)]));
}catch(error){
 const message=error instanceof Error?error.message:'办公文件解析失败';
 process.stdout.end(Buffer.concat([Buffer.from([1]),Buffer.from(message.slice(0,500))]));
}})();
