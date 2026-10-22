import {useRef,useState,useEffect} from 'react';
import {CloudError} from '../../../common/nbstore/src/cloud';
import {workbookSheets} from '../../../common/nbstore/src/office-workbook';
import {structureRevision,type SheetOperation} from '../../../common/nbstore/src/spreadsheet-operation';
import {ResourceDialog} from './dialogs';
import type {LocalWorkspace,OfficeSpreadsheetDocument} from './workspace';
export function WorksheetControls({item,workspace,busy,onBusy,onSelect}:{item:OfficeSpreadsheetDocument;workspace:LocalWorkspace;busy:boolean;onBusy:(busy:boolean)=>void;onSelect:(id:string)=>void}){
 const [dialog,setDialog]=useState<'addSheet'|'renameSheet'|'deleteSheet'>(),[error,setError]=useState(''),pending=useRef<SheetOperation>(),dialogTarget=useRef<{id:string;name:string}>();
 const sheets=workbookSheets(item.doc),selected=sheets.find(sheet=>sheet.id===item.sheetId)??sheets[0],index=sheets.indexOf(selected),readOnly=workspace.documentReadOnly(item.id);
 useEffect(()=>{if(!dialog)pending.current=undefined;},[selected.id,dialog]);
 const open=(type:'addSheet'|'renameSheet'|'deleteSheet')=>{pending.current=undefined;dialogTarget.current={id:selected.id,name:selected.name};setDialog(type);};
 const submit=async(type:SheetOperation['type'],name?:string,target=selected.id)=>{
  if(readOnly||busy)return;onBusy(true);setError('');try{
   await workspace.prepareOfficeBackup(item.id);
   const operation=pending.current?.type===type&&pending.current?.sheet===target&&pending.current?.name===name?pending.current:{type,sheet:target,name,operationId:crypto.randomUUID(),base:structureRevision(item.doc)};pending.current=operation;
   await workspace.submitSpreadsheetOperation({...item,sheetId:operation.sheet},operation);pending.current=undefined;onSelect(type==='addSheet'?'tab-'+operation.operationId:type==='deleteSheet'?workbookSheets(item.doc)[0].id:selected.id);
  }catch(e){if(e instanceof CloudError&&e.status===409)pending.current=undefined;await workspace.refreshSpreadsheet(item).catch(()=>{});throw e;}finally{onBusy(false);}
 };
 const move=(type:'moveSheetLeft'|'moveSheetRight')=>void submit(type).catch(error=>setError(error instanceof Error?error.message:'工作表操作失败'));
 return <><div className="worksheet-controls"><button disabled={readOnly||busy} onClick={()=>open('addSheet')}>新增工作表</button><button disabled={readOnly||busy} onClick={()=>open('renameSheet')}>重命名工作表</button><button disabled={readOnly||busy||sheets.length<=1} onClick={()=>open('deleteSheet')}>删除工作表</button><button disabled={readOnly||busy||index===0} onClick={()=>move('moveSheetLeft')}>左移工作表</button><button disabled={readOnly||busy||index===sheets.length-1} onClick={()=>move('moveSheetRight')}>右移工作表</button></div>{error&&<p role="alert">{error}</p>}
 {dialog&&<ResourceDialog title={dialog==='addSheet'?'新增工作表':dialog==='renameSheet'?'重命名工作表':'删除工作表'} nameLabel="工作表名称" initialName={dialog==='renameSheet'?dialogTarget.current!.name:''} confirm={dialog==='deleteSheet'} onClose={()=>setDialog(undefined)} action={async name=>{await submit(dialog,name||undefined,dialogTarget.current!.id);}}><p>{dialog==='deleteSheet'?`删除「${dialogTarget.current!.name}」及其中所有内容。若其他公式可能引用此表，会阻止删除。`:'名称最多 31 个字符且不能重复。已有公式可能引用此表时，会阻止改名，请先处理公式。'}</p></ResourceDialog>}
 </>;
}
