import {changeSheetStructure} from './office-sheet-structure';
import {sheetColumnIds,sheetRowIdentity} from '../../../../common/nbstore/src/sheet-columns';
import {formulaMayReferenceSheet} from '../../../../common/nbstore/src/spreadsheet-formula';
import {workbookSheet,workbookSheets,FIRST_WORKSHEET,MAX_WORKBOOK_SHEETS,validWorksheetName} from '../../../../common/nbstore/src/office-workbook';
import {spreadsheetCellType,setSpreadsheetCellType,clearRowCellTypes,clearSheetCellTypes,validSpreadsheetCellType} from '../../../../common/nbstore/src/spreadsheet-cell-type';
import * as Y from 'yjs';
import {createHash} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {cellRevision,structureRevision,MAX_SHEET_BATCH_CELLS,type SheetCellEdit,type SheetOperation} from '../../../../common/nbstore/src/spreadsheet-operation';
import {validateSpreadsheetUpdate} from './office-spreadsheet';
import {sheetCellKey,resolveSheetMerges,adjustSheetFormatsForStructure,clearSheetFormats,validSheetCellFormat,validSheetNumberFormat,SHEET_FORMAT_LIMITS,type SheetCellFormat} from '../../../../common/nbstore/src/spreadsheet-format';
import {adjustSheetObjectsForStructure,clearSheetObjects,resolveSheetRange,sheetObjectRevision,sheetObjects,validSheetObject,validSheetObjectId,SHEET_OBJECT_LIMITS,type SheetObject} from '../../../../common/nbstore/src/spreadsheet-objects';
export class SheetOperationError extends Error {constructor(message:string,readonly status=400){super(message);}}
export function submitSheetOperation(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:SheetOperation){
 if(!input||!['insertRow','deleteRow','insertColumn','deleteColumn','setCell','appendRow','deleteLastRow','setCells','addSheet','deleteSheet','moveSheetLeft','moveSheetRight','renameSheet','setFormat','mergeCells','unmergeCells','putSheetObject','deleteSheetObject'].includes(input.type)||typeof input.operationId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(input.operationId)||typeof input.base!=='string'||input.base.length>131072)throw new SheetOperationError('表格操作参数无效');
 if(input.cellType!==undefined&&(input.type!=='setCell'||!validSpreadsheetCellType(input.cellType)))throw new SheetOperationError('单元格内容类型无效');
 if(input.cells!==undefined&&input.type!=='setCells')throw new SheetOperationError('批量单元格参数无效');
 let edits:SheetCellEdit[]|undefined;
 if(input.type==='setCells'){
  if(!Array.isArray(input.cells)||!input.cells.length||input.cells.length>MAX_SHEET_BATCH_CELLS)throw new SheetOperationError('批量粘贴需包含 1–500 个单元格');
  const locations=new Set<string>();let length=0;
  edits=input.cells.map(cell=>{
   if(!cell||!Number.isSafeInteger(cell.row)||cell.row<0||!Number.isSafeInteger(cell.column)||cell.column<0||cell.column>=16384||typeof cell.value!=='string'||cell.value.length>10000||typeof cell.base!=='string'||cell.base.length>131072||cell.cellType!==undefined&&!validSpreadsheetCellType(cell.cellType))throw new SheetOperationError('批量单元格参数无效');
   const key=cell.row+':'+cell.column;if(locations.has(key))throw new SheetOperationError('批量粘贴包含重复单元格');locations.add(key);length+=cell.value.length;if(length>100000)throw new SheetOperationError('批量粘贴文字超过 10 万字符');
   return {row:cell.row,column:cell.column,value:cell.value,base:cell.base,...(cell.cellType!==undefined?{cellType:cell.cellType}:{})};
  });
 }
 if(input.sheet!==undefined&&(typeof input.sheet!=='string'||! /^[a-zA-Z0-9-]{1,128}$/.test(input.sheet)))throw new SheetOperationError('工作表编号无效');
 if(input.name!==undefined&&(!['addSheet','renameSheet'].includes(input.type)||!validWorksheetName(input.name)))throw new SheetOperationError('工作表名称无效，最多 31 个字符且不能包含特殊字符');
 const formatOperation=['setFormat','mergeCells','unmergeCells'].includes(input.type),stableKey=(value:unknown)=>typeof value==='string'&&/^[0-9]+-[0-9]+$/.test(value),stableColumn=(value:unknown)=>typeof value==='string'&&/^[a-zA-Z0-9-]{1,180}$/.test(value);
 if((input.targets!==undefined||input.format!==undefined)!==(input.type==='setFormat')||(input.rowSpan!==undefined||input.colSpan!==undefined)!==(input.type==='mergeCells'))throw new SheetOperationError('格式或合并参数无效');
 if(input.type==='setFormat'){
  const format=input.format,targets=input.targets;
  if(!Array.isArray(targets)||!targets.length||targets.length>SHEET_FORMAT_LIMITS.batch||targets.some(target=>!target||Object.keys(target).length!==2||!stableKey(target.rowKey)||!stableColumn(target.columnKey))||new Set(targets.map(target=>target.rowKey+':'+target.columnKey)).size!==targets.length)throw new SheetOperationError('格式目标需包含 1–2000 个不重复的单元格');
  if(!format||typeof format!=='object'||Array.isArray(format)||!Object.keys(format).length||Object.entries(format).some(([key,value])=>value!==null&&!(key==='numberFormat'?validSheetNumberFormat(value):validSheetCellFormat({[key]:value}))))throw new SheetOperationError('单元格格式无效');
 }
 if(input.type==='mergeCells'&&(!Number.isInteger(input.rowSpan)||!Number.isInteger(input.colSpan)||input.rowSpan!<1||input.colSpan!<1||input.rowSpan!>SHEET_FORMAT_LIMITS.span||input.colSpan!>SHEET_FORMAT_LIMITS.span||input.rowSpan!*input.colSpan!<2))throw new SheetOperationError('合并范围无效，至少包含两个单元格');
 if(formatOperation&&input.type!=='setFormat'&&(!stableKey(input.rowKey)||!stableColumn(input.columnKey)))throw new SheetOperationError('合并单元格稳定标识无效');
 if((input.rowKey!==undefined||input.columnKey!==undefined)&&!formatOperation&&(input.type!=='setCell'||typeof input.rowKey!=='string'||! /^[0-9]+-[0-9]+$/.test(input.rowKey)||typeof input.columnKey!=='string'||! /^[a-zA-Z0-9-]{1,180}$/.test(input.columnKey)))throw new SheetOperationError('单元格稳定标识无效');
 // Conditional formats and charts: `objectId` plus the whole new value; `base` is the stored object's revision.
 const objectOperation=input.type==='putSheetObject'||input.type==='deleteSheetObject';
 if((input.objectId!==undefined||input.object!==undefined)&&!objectOperation)throw new SheetOperationError('工作表对象参数无效');
 if(objectOperation&&(!validSheetObjectId(input.objectId)||(input.type==='putSheetObject')!==(input.object!==undefined)||input.type==='putSheetObject'&&!validSheetObject(input.object)||input.type==='putSheetObject'&&input.object!.sheet!==(input.sheet??FIRST_WORKSHEET)))throw new SheetOperationError('条件格式或图表参数无效');
 const sheetId=input.sheet??FIRST_WORKSHEET;
 const fields:unknown[]=[user,input.type,input.row??null,input.column??null,input.value??null,input.base];if(input.cellType!==undefined)fields.push(input.cellType);if(edits)fields.push(edits);if(input.sheet!==undefined)fields.push(input.sheet);if(input.name!==undefined)fields.push(input.name);
 if(input.objectId!==undefined)fields.push(input.objectId,input.object??null);if(input.rowKey!==undefined)fields.push(input.rowKey,input.columnKey);if(input.targets!==undefined)fields.push(input.targets,input.format);if(input.rowSpan!==undefined)fields.push(input.rowSpan,input.colSpan);
 const signature=createHash('sha256').update(JSON.stringify(fields)).digest('hex');
 return database.db.transaction(()=>{
  const scope='cloud:'+workspace,current=database.load(scope,id);if(!current)throw new SheetOperationError('表格不存在',404);
  const doc=new Y.Doc();
  try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('sheetOperations');
   const previous=receipts.get(input.operationId);
   if(previous){if(previous!==signature)throw new SheetOperationError('操作编号已用于其他提交',409);return {update:current,repeated:true,before:'',after:''};}
   if(receipts.size>=10000)throw new SheetOperationError('本表操作记录已达到首版上限，请导出后创建新副本',409);
   const before=structureRevision(doc);let sheet:Y.Array<Y.Array<string>>;try{sheet=workbookSheet(doc,sheetId).sheet;}catch{throw new SheetOperationError('工作表不存在或已删除',404);}
   const vector=Y.encodeStateVector(doc);
   if(['addSheet','deleteSheet','moveSheetLeft','moveSheetRight','renameSheet'].includes(input.type)){
    if(structureRevision(doc)!==input.base)throw new SheetOperationError('工作簿已更新，请核对最新工作表后重试',409);
    const all=workbookSheets(doc),active=all.find(sheet=>sheet.id===sheetId)!;
    if(input.type==='addSheet'&&all.length>=MAX_WORKBOOK_SHEETS)throw new SheetOperationError('工作表最多 100 张');
    if(['addSheet','renameSheet'].includes(input.type)&&(!validWorksheetName(input.name)||all.some(sheet=>sheet.id!== (input.type==='renameSheet'?sheetId:'')&&sheet.name.toLowerCase()===input.name!.toLowerCase())))throw new SheetOperationError('工作表名称无效或已存在');
    if(input.type==='deleteSheet'&&all.length<=1)throw new SheetOperationError('至少保留一张工作表');
    let referenceBudget=0;
    if(input.type==='deleteSheet'||input.type==='renameSheet')for(const other of all){if(input.type==='deleteSheet'&&other.id===sheetId)continue;for(const [r,row]of other.sheet.toArray().entries())for(const [c,value]of row.toArray().entries()){if(!value.trim().startsWith('=')||spreadsheetCellType(doc,r,c,other.id)==='text')continue;referenceBudget+=value.length;if(referenceBudget>200000)throw new SheetOperationError('工作表引用检查超过配额，请先简化公式或保留工作表');if(formulaMayReferenceSheet(value,active.name))throw new SheetOperationError('仍有公式可能引用此工作表，请先处理相关公式再删除或改名');}}
    const entries=doc.getMap<Y.Map<unknown>>('workbookSheets'),order=doc.getArray<string>('sheetOrder');
    if(!order.length){const first=new Y.Map<unknown>();first.set('name',active.name);first.set('defaultType',active.defaultType);entries.set(FIRST_WORKSHEET,first);order.push([FIRST_WORKSHEET]);}
    const index=order.toArray().indexOf(sheetId);
    if(input.type==='addSheet'){const grid=new Y.Array<Y.Array<string>>(),meta=new Y.Map<unknown>();meta.set('name',input.name);meta.set('defaultType','auto');meta.set('sheet',grid);const id='tab-'+input.operationId;entries.set(id,meta);order.push([id]);for(let r=0;r<20;r++){const row=new Y.Array<string>();row.push(Array(8).fill(''));grid.push([row]);}}
    else if(input.type==='renameSheet')entries.get(sheetId)!.set('name',input.name);
    else if(input.type==='deleteSheet'){clearSheetCellTypes(doc,sheetId);clearSheetFormats(doc,sheetId);clearSheetObjects(doc,sheetId);if(doc.share.has('sheetColumns'))doc.getMap('sheetColumns').delete(sheetId);order.delete(index,1);entries.delete(sheetId);if(sheetId===FIRST_WORKSHEET)sheet.delete(0,sheet.length);}
    else{const target=input.type==='moveSheetLeft'?index-1:index+1;if(target>=0&&target<order.length){order.delete(index,1);order.insert(target,[sheetId]);}}
   }else if(objectOperation){
    const objects=doc.getMap<SheetObject>('sheetObjects'),current=objects.get(input.objectId!);
    if(sheetObjectRevision(validSheetObject(current)?current:undefined)!==input.base)throw new SheetOperationError(current?'此条件格式或图表已被其他成员修改或删除，请刷新后重试':'对象已被删除，请刷新后重试',409);
    if(current&&validSheetObject(current)&&current.sheet!==sheetId)throw new SheetOperationError('对象不属于此工作表',409);
    if(input.type==='deleteSheetObject'){if(!current)throw new SheetOperationError('对象已被删除',409);objects.delete(input.objectId!);}
    else{
     const value=input.object!;if(!resolveSheetRange(sheet,value.range))throw new SheetOperationError('所选区域的行列已删除，请重新选择',409);
     if(value.kind==='chart'&&!resolveSheetRange(sheet,{top:value.anchor.row,bottom:value.anchor.row,left:value.anchor.column,right:value.anchor.column}))throw new SheetOperationError('图表位置的行列已删除，请重新放置',409);
     if(!current){const same=sheetObjects(doc,sheetId).filter(([,object])=>object.kind===value.kind).length;if(objects.size>=SHEET_OBJECT_LIMITS.objects||same>=(value.kind==='chart'?SHEET_OBJECT_LIMITS.chartsPerSheet:SHEET_OBJECT_LIMITS.rulesPerSheet))throw new SheetOperationError(value.kind==='chart'?'每张工作表最多 50 个图表':'每张工作表最多 200 条条件格式规则',413);}
     objects.set(input.objectId!,value);
    }
   }else if(formatOperation){
    const rows=sheet.toArray(),columns=sheetColumnIds(sheet),rowIndex=new Map(rows.map((row,i)=>[sheetRowIdentity(row),i])),locate=(rowKey:string,columnKey:string)=>{const row=rowIndex.get(rowKey),column=columns.indexOf(columnKey);if(row===undefined||column<0)throw new SheetOperationError('目标单元格所属行列已删除，请重新选择',409);return {row,column};};
    if(input.type==='setFormat'){
     // Formatting is commutative per field: each patch applies in commit order without discarding other fields.
     const formats=doc.getMap<SheetCellFormat>('sheetFormats');
     for(const target of input.targets!){locate(target.rowKey,target.columnKey);const key=sheetCellKey(sheetId,target.rowKey,target.columnKey),next:Record<string,unknown>={...formats.get(key)};for(const [field,value]of Object.entries(input.format!)){if(value===null)delete next[field];else next[field]=value;}if(Object.keys(next).length)formats.set(key,next as SheetCellFormat);else formats.delete(key);}
     if(formats.size>SHEET_FORMAT_LIMITS.formats)throw new SheetOperationError('工作簿格式单元格超过 5 万，请先清除部分格式',413);
    }else{
     const anchor=locate(input.rowKey!,input.columnKey!),key=sheetCellKey(sheetId,input.rowKey!,input.columnKey!),merges=doc.getMap<{rowSpan:number;colSpan:number}>('sheetMerges'),layout=resolveSheetMerges(doc,sheetId,sheet);
     if(input.type==='unmergeCells'){if(!layout.anchors.has(anchor.row+':'+anchor.column))throw new SheetOperationError('此单元格未合并或合并已失效',409);merges.delete(layout.anchors.get(anchor.row+':'+anchor.column)!.key);}
     else{
      if(anchor.row+input.rowSpan!>rows.length||anchor.column+input.colSpan!>columns.length)throw new SheetOperationError('合并范围超出工作表');
      for(let r=anchor.row;r<anchor.row+input.rowSpan!;r++)for(let c=anchor.column;c<anchor.column+input.colSpan!;c++){
       if(layout.owner.has(r+':'+c))throw new SheetOperationError('合并范围与已合并单元格重叠，请先拆分',409);
       if((r!==anchor.row||c!==anchor.column)&&(rows[r].get(c)??'')!=='')throw new SheetOperationError('只能合并除左上角外均为空的区域，请先移走其他单元格内容',409);
      }
      merges.set(key,{rowSpan:input.rowSpan!,colSpan:input.colSpan!});if(merges.size>SHEET_FORMAT_LIMITS.merges)throw new SheetOperationError('合并区域超过 1 万个',413);
     }
    }
   }else if(input.type==='setCells'){
    // Validate every target against the same persisted state before applying any edit.
    const covered=resolveSheetMerges(doc,sheetId,sheet).owner;
    for(const cell of edits!){
     if(cell.row>=sheet.length||cell.column>=sheet.get(cell.row).length)throw new SheetOperationError('批量粘贴超出现有网格，请先调整粘贴范围');
     if(cell.value&&covered.has(cell.row+':'+cell.column)&&covered.get(cell.row+':'+cell.column)!==cell.row+':'+cell.column)throw new SheetOperationError(`第 ${cell.row+1} 行第 ${cell.column+1} 列位于合并区域内，整批未写入，请先拆分`,409);
     if(cellRevision(sheet,cell.row,cell.column)!==cell.base)throw new SheetOperationError(`第 ${cell.row+1} 行第 ${cell.column+1} 列已更新，整批未写入，请重新预览`,409);
     const kind=cell.cellType??spreadsheetCellType(doc,cell.row,cell.column,sheetId);if(kind==='formula'&&!cell.value.trim().startsWith('='))throw new SheetOperationError('公式类型的内容必须以 = 开头，请选择文本或自动');
    }
    for(const cell of edits!){const target=sheet.get(cell.row);target.delete(cell.column,1);target.insert(cell.column,[cell.value]);if(cell.cellType!==undefined)setSpreadsheetCellType(doc,cell.row,cell.column,cell.cellType,sheetId);}
   }else if(input.type==='setCell'){
    const {value}=input;let {row,column}=input;if(input.rowKey!==undefined){row=sheet.toArray().findIndex(row=>sheetRowIdentity(row)===input.rowKey);column=sheetColumnIds(sheet).indexOf(input.columnKey!);if(row<0||column<0)throw new SheetOperationError('目标单元格所属行列已删除，草稿已保留',409);}
    if(typeof row!=='number'||!Number.isSafeInteger(row)||row<0||row>=sheet.length||typeof column!=='number'||!Number.isSafeInteger(column)||column<0||column>=16384||typeof value!=='string'||value.length>10000)throw new SheetOperationError('单元格参数无效');
    if(cellRevision(sheet,row,column)!==input.base)throw new SheetOperationError('此单元格已被他人修改，请选择保留的内容',409);
    const owner=resolveSheetMerges(doc,sheetId,sheet).owner.get(row+':'+column);if(value&&owner&&owner!==row+':'+column)throw new SheetOperationError('此单元格已被合并，输入已保留，请先拆分合并区域',409);
    const kind=input.cellType??spreadsheetCellType(doc,row,column,sheetId);if(kind==='formula'&&!value.trim().startsWith('='))throw new SheetOperationError('公式类型的内容必须以 = 开头，请选择文本或自动');
    const target=sheet.get(row);
    while(target.length<=column)target.push(['']);target.delete(column,1);target.insert(column,[value]);if(input.cellType!==undefined)setSpreadsheetCellType(doc,row,column,kind,sheetId);
   }else{
    if(structureRevision(doc)!==input.base)throw new SheetOperationError('表格已更新，请刷新后重试结构操作',409);
        const action=input.type==='appendRow'?'insertRow':input.type==='deleteLastRow'?'deleteRow':input.type,index=input.type==='appendRow'?sheet.length:input.type==='deleteLastRow'?sheet.length-1:action.endsWith('Column')?input.column:input.row;
    try{if(Number.isSafeInteger(index)&&index!>=0&&index!<=(action.endsWith('Column')?sheetColumnIds(sheet).length:sheet.length)){adjustSheetFormatsForStructure(doc,sheetId,sheet,action.endsWith('Column')?'column':'row',index!,action.startsWith('delete'));adjustSheetObjectsForStructure(doc,sheetId,sheet,action.endsWith('Column')?'column':'row',index!,action.startsWith('delete'));}changeSheetStructure(doc,sheetId,action,index!,input.operationId);}catch(error){throw new SheetOperationError(error instanceof Error?error.message:'结构操作失败');}
   }
   receipts.set(input.operationId,signature);
   const update=Y.encodeStateAsUpdate(doc,vector);validateSpreadsheetUpdate(current,update);
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(doc)),scope,id);
   database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);
   // Revisions around this commit let an offline client rebase its next queued edit when nothing else changed.
   return {update,repeated:false,before,after:structureRevision(doc)};
  }finally{doc.destroy();}
 }).immediate();
}
