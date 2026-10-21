import {unpackOfficeMedia,checkOfficeMedia} from './office-media';
import * as Y from 'yjs';
import {randomUUID} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {MAX_OFFICE_FILE_BYTES,officeKindFromFilename,prepareOfficeTextImport,type TextEncoding,type TextLineEnding} from '../../../../common/nbstore/src/office';
import {validateOfficeTextUpdate} from './office-text';
import {emptyPdfUpdate} from './office-pdf';
import {OfficeParserError} from './office-parser-process';
import type {OfficeParserFormat} from './office-parser-model';

export class OfficeTextImportError extends Error {}
export class OfficeImportError extends Error {constructor(message:string,readonly status=400){super(message);}}
export function importOfficeText(database:SQLiteProjectDatabase,workspace:string,user:string,name:string,bytes:Uint8Array,encoding?:TextEncoding,lineEnding?:TextLineEnding){
  let imported:ReturnType<typeof prepareOfficeTextImport>;
  try{imported=prepareOfficeTextImport(bytes,encoding,lineEnding);}catch(error){throw new OfficeTextImportError(error instanceof Error?error.message:'无法解析文本，请确认编码');}
  const doc=new Y.Doc();
  const id='doc:'+randomUUID(),assetId=randomUUID(),generation=randomUUID(),now=Date.now(),scope='cloud:'+workspace;
  try{
    doc.getText('text').insert(0,imported.text);
    const settings=doc.getMap('textSettings');for(const [key,value]of Object.entries(imported.settings))settings.set(key,value);
    const update=Y.encodeStateAsUpdate(doc);validateOfficeTextUpdate(null,update);
    // AttachmentFiles writes immutable bytes before the SQLite references. If the
    // reference transaction fails, the file is intentionally left unreferenced and
    // can be reclaimed later; no existing attachment can be damaged.
    database.attachments.put(scope,workspace,assetId,name,'text/plain',bytes);
    database.db.transaction(()=>{
      database.db.prepare('INSERT INTO cloud_office_resources(workspace,id,kind,name,mime,source_format,original_asset_id,model_version,generation,created_by,created,updated,deleted) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,0)').run(workspace,id,'text',name,'text/plain','txt',assetId,1,generation,user,now,now);
      database.db.prepare('INSERT INTO cloud_asset_owners(workspace,id,doc,user) VALUES (?,?,?,?)').run(workspace,assetId,id,user);
      database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run(scope,id,Buffer.from(update));
    }).immediate();
    return {id,kind:'text' as const,name,mime:'text/plain',sourceFormat:'txt',originalAssetId:assetId,modelVersion:1,generation,createdBy:user,created:now,updated:now,deleted:0};
  }finally{doc.destroy();}
}

/** Parse and validate before touching attachment references or taking the write lock. */
export async function importOfficeOriginal(database:SQLiteProjectDatabase,workspace:string,user:string,name:string,bytes:Uint8Array,parse:(format:OfficeParserFormat,bytes:Uint8Array)=>Promise<Uint8Array>,beforeCommit:()=>void=()=>{}){
  if(bytes.length>MAX_OFFICE_FILE_BYTES)throw new OfficeImportError('原件超过 20 MB');
  const format=officeKindFromFilename(name);if(!format)throw new OfficeImportError('暂不支持该办公文件类型');
  if(format.kind==='text')throw new OfficeImportError('TXT 文件请使用文本导入向导');
  if(format.kind==='pdf'&&(bytes.length<5||new TextDecoder().decode(bytes.subarray(0,5))!=='%PDF-'))throw new OfficeImportError('PDF 文件头无效');
  if(['word','presentation'].includes(format.kind)&&!(bytes[0]===0x50&&bytes[1]===0x4b))throw new OfficeImportError('Office 文件不是有效的 ZIP 包');
  if(['xlsx','xlsm','xltx','xltm'].includes(format.sourceFormat)&&!(bytes[0]===0x50&&bytes[1]===0x4b))throw new OfficeImportError('Excel 文件不是有效的 ZIP 包');
  let initial:Uint8Array;
  try{
    if(format.kind==='pdf'){const parsed=emptyPdfUpdate();try{initial=parsed.update;}finally{parsed.doc.destroy();}}
    else initial=await parse(format.kind==='word'?'docx':format.kind==='presentation'?'pptx':['csv','tsv'].includes(format.sourceFormat)?format.sourceFormat as 'csv'|'tsv':'xlsx',bytes);
  }catch(error){throw new OfficeImportError(error instanceof Error?error.message:'办公文件解析失败',error instanceof OfficeParserError?error.status:400);}
  const payload=unpackOfficeMedia(initial);checkOfficeMedia(payload.state,format.kind,payload.media);initial=payload.state;
  beforeCommit();
  const scope='cloud:'+workspace,id='doc:'+randomUUID(),assetId=randomUUID(),generation=randomUUID(),now=Date.now();
  database.attachments.put(scope,workspace,assetId,name,format.mime,bytes);
  database.db.transaction(()=>{
    database.db.prepare('INSERT INTO cloud_office_resources(workspace,id,kind,name,mime,source_format,original_asset_id,model_version,generation,created_by,created,updated,deleted) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,0)').run(workspace,id,format.kind,name,format.mime,format.sourceFormat,assetId,1,generation,user,now,now);
    database.db.prepare('INSERT INTO cloud_asset_owners(workspace,id,doc,user) VALUES (?,?,?,?)').run(workspace,assetId,id,user);
    for(const media of payload.media){const data=Buffer.from(media.data,'base64');database.attachments.put(scope,workspace,media.id,media.name,media.mime,data);database.db.prepare('INSERT INTO cloud_asset_owners(workspace,id,doc,user) VALUES (?,?,?,?)').run(workspace,media.id,id,user);}
    if(initial)database.db.prepare('INSERT INTO documents(scope,id,data) VALUES (?,?,?)').run(scope,id,Buffer.from(initial));
  }).immediate();
  return {id,kind:format.kind,name,mime:format.mime,sourceFormat:format.sourceFormat,originalAssetId:assetId,modelVersion:1,generation,createdBy:user,created:now,updated:now,deleted:0};
}
