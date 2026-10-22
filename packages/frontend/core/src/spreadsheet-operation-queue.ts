import type * as Y from 'yjs';
import {CloudError} from '../../../common/nbstore/src/cloud';
import type {SheetOperation} from '../../../common/nbstore/src/spreadsheet-operation';
/** Prefix of the error recorded on a cell draft whose submit never reached the server; such drafts retry automatically. */
export const OFFLINE_DRAFT='离线未提交';
export const sheetQueuePrefix=(id:string)=>'sheet-operation-queue:'+id+':';
export type QueuedSheetOperation={operation:SheetOperation;created:number;error?:string;conflict?:boolean};
const STRUCTURAL=new Set(['insertRow','deleteRow','insertColumn','deleteColumn','appendRow','deleteLastRow','addSheet','deleteSheet','moveSheetLeft','moveSheetRight','renameSheet']);
export const isStructuralSheetOperation=(operation:SheetOperation)=>STRUCTURAL.has(operation.type);
/** A failed fetch (offline, DNS, connection reset) rather than a server answer. */
export function isNetworkFailure(error:unknown){return error instanceof TypeError&&/fetch|network|load failed|connection/i.test(error.message);}
/**
 * Semantic operations that could not reach the server are kept per account in SQLite preferences and
 * replayed in creation order with their original ids. Every operation carries its own revision base, so a
 * replay either applies exactly what the author saw or stops with a conflict; nothing is applied optimistically.
 */
export class SpreadsheetOperationQueue{
 constructor(private readonly preferences:Y.Map<unknown>,private readonly resource:string){}
 private get prefix(){return sheetQueuePrefix(this.resource);}
 entries(){return [...this.preferences.entries()].filter(([key])=>key.startsWith(this.prefix)).map(([key,value])=>[key,value as QueuedSheetOperation] as const).sort(([a],[b])=>a.localeCompare(b));}
 get size(){return this.entries().length;}
 hasStructural(){return this.entries().some(([,entry])=>isStructuralSheetOperation(entry.operation));}
 add(operation:SheetOperation){const created=Date.now(),key=this.prefix+String(created).padStart(15,'0')+':'+operation.operationId;this.preferences.set(key,JSON.parse(JSON.stringify({operation,created})));return key;}
 update(key:string,patch:Partial<QueuedSheetOperation>){const current=this.preferences.get(key) as QueuedSheetOperation|undefined;if(current)this.preferences.set(key,JSON.parse(JSON.stringify({...current,...patch})));}
 delete(key:string){this.preferences.delete(key);}
 /**
  * Submit queued entries in order; stop at the first that is still offline, conflicts or fails.
  * A structure edit based on the state just before an earlier replayed edit is rebased onto that edit's
  * result, since the only change in between is the author's own queued work.
  */
 async drain(send:(operation:SheetOperation)=>Promise<{before:string;after:string}|void>){
  let replayed:{before:string;after:string}|undefined;
  for(const [key,entry]of this.entries()){
   if(entry.conflict||entry.error)return false;
   let operation=entry.operation;if(replayed&&replayed.before&&isStructuralSheetOperation(operation)&&operation.base===replayed.before){operation={...operation,base:replayed.after};this.update(key,{operation});}
   try{replayed=await send(operation)??undefined;this.delete(key);}
   catch(error){if(isNetworkFailure(error))return false;this.update(key,{error:error instanceof Error?error.message:'提交失败',conflict:error instanceof CloudError&&error.status===409});return false;}
  }
  return true;
 }
}
