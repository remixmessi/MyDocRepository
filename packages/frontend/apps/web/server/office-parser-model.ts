import {packOfficeMedia} from './office-media';
import {MAX_OFFICE_FILE_BYTES} from '../../../../common/nbstore/src/office';
import {encodeDelimitedWorkbook,encodeXlsxWorkbook,validateSpreadsheetUpdate} from './office-spreadsheet';
import {encodeDocxWord,validateWordUpdate} from './office-word';
import {encodePptxPresentation,validatePresentationUpdate} from './office-presentation';
export type OfficeParserFormat='docx'|'xlsx'|'pptx'|'csv'|'tsv';
export const OFFICE_PARSER_OUTPUT_BYTES=16*1024*1024;
export function parseOfficeModel(format:OfficeParserFormat,bytes:Uint8Array){
 if(bytes.length>MAX_OFFICE_FILE_BYTES)throw new Error('原件超过 20 MB');
 const parsed=format==='docx'?encodeDocxWord(bytes):format==='pptx'?encodePptxPresentation(bytes):format==='xlsx'?encodeXlsxWorkbook(bytes):encodeDelimitedWorkbook(new TextDecoder('utf-8',{fatal:true}).decode(bytes),format==='csv'?',':'\t');
 try{
  (format==='docx'?validateWordUpdate:format==='pptx'?validatePresentationUpdate:validateSpreadsheetUpdate)(null,parsed.update);
  if(parsed.update.length>OFFICE_PARSER_OUTPUT_BYTES)throw new Error('协作模型超过 16 MB');
  return packOfficeMedia(parsed.update,'media' in parsed?parsed.media as import('./office-media').OfficeMedia[]:[]);
 }finally{parsed.doc.destroy();}
}
