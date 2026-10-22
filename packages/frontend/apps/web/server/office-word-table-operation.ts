import * as Y from 'yjs';
import {createHash} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {applyWordTableOperation,findWordTable,validWordTableOperation,wordTableRevision,type WordTableOperation} from '../../../../common/nbstore/src/office-word-table-edit';
import {validateWordUpdate} from './office-word';
export class WordTableOperationError extends Error{constructor(message:string,readonly status=400){super(message);}}
const ORIGIN=Symbol('word-table-operation');

/**
 * Team workspaces change Word table structure only here: the table's structure revision must match what the user saw,
 * the change runs on the saved state, and a receipt makes a retried request (lost response) return the same result.
 * Concurrent cell text edits keep merging through Yjs; only row/column/merge changes are serialized.
 */
export function submitWordTableOperation(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:WordTableOperation):{update:Uint8Array;focus?:string;repeated:boolean}{
 if(!validWordTableOperation(input))throw new WordTableOperationError('表格操作参数无效');
 const signature=createHash('sha256').update(JSON.stringify([user,input.table,input.base,input.action,input.cell??null])).digest('hex');
 return database.db.transaction(()=>{
  const scope='cloud:'+workspace,current=database.load(scope,id);if(!current)throw new WordTableOperationError('文档不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('wordTableOperations'),previous=receipts.get(input.operationId);
   if(previous){const [saved,focus]=previous.split(':');if(saved!==signature)throw new WordTableOperationError('操作编号已用于其他提交',409);return {update:current,...(focus?{focus}:{}),repeated:true};}
   if(receipts.size>=10000)throw new WordTableOperationError('表格操作记录已达上限，请备份后恢复为新文件',409);
   const table=findWordTable(doc,input.table);if(!table)throw new WordTableOperationError('表格已被其他成员删除',409);
   if(wordTableRevision(table)!==input.base)throw new WordTableOperationError('表格结构已被其他成员修改，请核对后重试',409);
   let focus:string|undefined;try{focus=applyWordTableOperation(doc,input,ORIGIN);}catch(error){throw new WordTableOperationError(error instanceof Error?error.message:'表格操作失败',409);}
   receipts.set(input.operationId,signature+':'+(focus??''));const update=Y.encodeStateAsUpdate(doc);
   try{validateWordUpdate(null,update);}catch(error){throw new WordTableOperationError(error instanceof Error?error.message:'Word 模型无效');}
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(update),scope,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);
   return {update,...(focus?{focus}:{}),repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}
