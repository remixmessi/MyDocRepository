import {OfficeHistoryDialog,DeletedOfficeHistoryDialog} from './office-history-ui';
import {OfficeBackupDialog} from './office-backup-ui';
import {OfficeExportHistoryDialog} from './office-export-history-ui';
import {useEffect,useState} from 'react';
import type {LocalWorkspace} from './workspace';
import {TextImportDialog} from './text-import-ui';
import {OfficeImportDialog} from './office-import-ui';

export function OfficeFileDialogs({workspace,onSelect}:{workspace:LocalWorkspace;onSelect:(id:string)=>void}){
  const [error,setError]=useState('');
  const [folderId,setFolderId]=useState<string>();
  const finish=async(id:string,close:()=>void)=>{try{if(folderId&&workspace.folders.has(folderId)){workspace.moveDocument(id,folderId);await workspace.persistence.flush();}}finally{close();onSelect(id);}};
  const [backup,setBackup]=useState<{id?:string}>();
  const [history,setHistory]=useState<{id:string;name?:string;deleted?:boolean}>(),[deletedHistory,setDeletedHistory]=useState(false);
  const [importOpen,setImportOpen]=useState(false),[exportsFor,setExportsFor]=useState<string>();
  const [officeImportOpen,setOfficeImportOpen]=useState(false);
  useEffect(()=>{const handle=(event:Event)=>{const detail=(event as CustomEvent<{workspace:string;kind:string;id?:string;folderId?:string}>).detail;if(detail.workspace!==workspace.id)return;setError('');setFolderId(detail.folderId);switch(detail.kind){case 'history':if(detail.id)setHistory({id:detail.id});break;case 'backup':setBackup({id:detail.id});break;case 'exports':if(detail.id)setExportsFor(detail.id);break;case 'deleted':setDeletedHistory(true);break;case 'text':setImportOpen(true);break;case 'import':setOfficeImportOpen(true);break;}};window.addEventListener('office-file-action',handle);return()=>window.removeEventListener('office-file-action',handle);},[workspace]);
  return <>{error&&<p role="alert">{error}</p>}
    {deletedHistory&&<DeletedOfficeHistoryDialog workspace={workspace} onClose={()=>setDeletedHistory(false)} onSelect={file=>{setDeletedHistory(false);setHistory({...file,deleted:true});}}/>}
    {history&&<OfficeHistoryDialog workspace={workspace} id={history.id} title={history.name} deleted={history.deleted} onClose={()=>setHistory(undefined)} onFinish={id=>{setHistory(undefined);onSelect(id);}}/>}
    {exportsFor&&<OfficeExportHistoryDialog workspace={workspace} id={exportsFor} onClose={()=>setExportsFor(undefined)}/>}
    {backup&&<OfficeBackupDialog workspace={workspace} id={backup.id} onClose={()=>setBackup(undefined)} onFinish={id=>{setBackup(undefined);onSelect(id);}}/>}
    {importOpen&&<TextImportDialog workspace={workspace} onClose={()=>setImportOpen(false)} onFinish={id=>void finish(id,()=>setImportOpen(false)).catch(error=>setError(error instanceof Error?error.message:'文件移动失败'))}/>}
    {officeImportOpen&&<OfficeImportDialog workspace={workspace} onClose={()=>setOfficeImportOpen(false)} onFinish={id=>void finish(id,()=>setOfficeImportOpen(false)).catch(error=>setError(error instanceof Error?error.message:'文件移动失败'))}/>}
  </>;
}
