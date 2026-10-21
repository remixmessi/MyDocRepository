import {unpackOfficeMedia,checkOfficeMedia} from './office-media';
import {presentationSlides} from '../../../../common/nbstore/src/office-presentation';
import * as Y from 'yjs';
import type {OfficeKind} from '../../../../common/nbstore/src/office';
import {validateOfficeTextUpdate} from './office-text';
import {validateWordUpdate} from './office-word';
import {validateSpreadsheetUpdate} from './office-spreadsheet';
import {validatePdfUpdate} from './office-pdf';
import {validatePresentationUpdate} from './office-presentation';

export function validateOfficeBackupModel(kind:OfficeKind,bytes:Uint8Array){
 const payload=unpackOfficeMedia(bytes);bytes=payload.state;
 const validate={text:validateOfficeTextUpdate,word:validateWordUpdate,spreadsheet:validateSpreadsheetUpdate,pdf:validatePdfUpdate,presentation:validatePresentationUpdate}[kind];
 if(!validate)throw new Error('备份模型类型无效');validate(null,bytes);checkOfficeMedia(bytes,kind,payload.media);
 const doc=new Y.Doc();try{
  Y.applyUpdate(doc,bytes);if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('备份模型缺少依赖');

  if(doc.share.has('officeCommentOperations'))doc.getMap('officeCommentOperations').clear();
  // Operation receipts belong to the old resource and must not exhaust the new one.
  if(kind==='pdf'){doc.getMap('pdfOperations').clear();doc.getMap('pdfOperations').set('$schema','0'.repeat(64));}
  if(kind==='spreadsheet')doc.getMap('sheetOperations').clear();
  if(kind==='presentation'&&doc.share.has('presentationOperations'))doc.getMap('presentationOperations').clear();
  if(kind==='word'&&doc.share.has('wordTableOperations'))doc.getMap('wordTableOperations').clear();
  return Y.encodeStateAsUpdate(doc);
 }finally{doc.destroy();}
}
