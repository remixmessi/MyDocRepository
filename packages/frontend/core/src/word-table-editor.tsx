import {WordEditingContext} from './word-editing-context';
import {useContext,useRef,useState} from 'react';
import {Y} from '@mydoc/crdt';
import {wordTableRows,wordTableLayout,WORD_TABLE_LIMITS} from '../../../common/nbstore/src/office-word-table';
import {appendWordTableRow,insertWordTableRow,deleteWordTableRow,insertWordTableColumn,deleteWordTableColumn,mergeWordTableCell,splitWordTableCell} from './office-word-table';
import {WordParagraphEditor} from './word-paragraph-editor';
import {findWordTableCell,wordItemIdentity,wordTableRevision,type WordTableAction,type WordTableOperation} from '../../../common/nbstore/src/office-word-table-edit';
/** Team workspaces send table structure changes to the server; resolves to the identity of the cell to focus. */
export type WordTableRemote=(operation:WordTableOperation)=>Promise<string|undefined>;
function WordTableCell({text,label,readOnly,assertWritable,marker}:{text:Y.Text;label:string;readOnly:boolean;assertWritable:()=>void;marker:string}){
 return <WordParagraphEditor paragraph={text} label={label} inputLabel={label} marker={marker} readOnly={readOnly} assertWritable={assertWritable} cell/>;
}
/**
 * One floating "表格工具" bar acting on the current cell, grouped like Word's table layout tab. Row tools no longer live
 * in an extra cell of the focused row, which used to add a column to that row only and break the fixed table layout.
 */
export function WordTableEditor({table,index,readOnly,assertWritable,onDelete,markers,remote}:{table:Y.Map<unknown>;index:number;readOnly:boolean;assertWritable:()=>void;onDelete:()=>void;markers:Map<Y.Text,string>;remote?:WordTableRemote}){
 const [error,setError]=useState(''),keys=useRef(new WeakMap<object,string>());const rows=wordTableRows(table),layout=wordTableLayout(table),editing=useContext(WordEditingContext),cell=[...layout.anchors.keys()].find(text=>text===editing.active),span=cell&&layout.anchors.get(cell);
 const key=(object:object)=>{let value=keys.current.get(object);if(!value){value=crypto.randomUUID();keys.current.set(object,value);}return value;};
 const act=(action:()=>void)=>{try{assertWritable();action();setError('');}catch(error){setError(error instanceof Error?error.message:'表格更新失败');}};
 // A failed request keeps its operation id so "重试" cannot apply it twice; a conflict drops it after reloading.
 const [busy,setBusy]=useState(false),pending=useRef<WordTableOperation>(),activeRef=useRef(editing.active);activeRef.current=editing.active;
 // The caret only follows the result if the user is still in the cell they acted on; never pull it back from elsewhere.
 const send=async(operation:WordTableOperation)=>{if(!remote)return;setBusy(true);setError('');pending.current=operation;const from=editing.active;
  try{assertWritable();const focus=await remote(operation);pending.current=undefined;const target=focus?findWordTableCell(table,focus)?.cell:undefined;if(target&&activeRef.current===from)editing.focus(target);}
  catch(error){if((error as {status?:number}).status===409)pending.current=undefined;setError(error instanceof Error?error.message:'表格更新失败');}finally{setBusy(false);}};
 /** Personal workspaces run `local` (instant, undoable); team workspaces send the same change as a semantic operation. */
 const change=(action:WordTableAction,local:()=>void)=>{if(!remote){act(local);return;}const identity=wordItemIdentity(table);if(!identity)return;void send({operationId:crypto.randomUUID(),table:identity,base:wordTableRevision(table),action,...(cell&&wordItemIdentity(cell)?{cell:wordItemIdentity(cell)}:{})});};
 const columns=Number(table.get('columns')),list=rows.toArray(),full=rows.length>=WORD_TABLE_LIMITS.rows;
 // A merged cell covers several rows: insert above its first row and below its last one.
 const top=span?list[span.row]:undefined,bottom=span?list[span.row+span.rowSpan-1]:undefined,rowLabel=span?`表格 ${index+1} 第 ${span.row+1} 行`:'当前行';
 const keep=(event:React.MouseEvent)=>event.preventDefault();
 // Like Word, the caret moves into the new row; this also keeps the toolbar (shown on focus) available.
 const focusRow=(row:Y.Array<Y.Text>)=>{const target=row.get(Math.min(span?.column??0,row.length-1));if(target)editing.focus(target);};
 const removeRow=()=>act(()=>{const at=span!.row;deleteWordTableRow(table,top!);const next=rows.get(Math.min(at,rows.length-1));const target=next?.get(Math.min(span!.column,next.length-1));if(target)editing.focus(target);});
 return <section className="office-word-table" tabIndex={0} aria-label={`第 ${index+1} 个表格`}><header role="toolbar" aria-label={`表格 ${index+1} 工具`}>
   <span className="word-table-size">表格 · {rows.length} 行 × {columns} 列</span>
   <span className="word-table-group" role="group" aria-label="行"><button disabled={busy||readOnly||!top||full} aria-label={`在${rowLabel}上方插入`} title="在上方插入行" onMouseDown={keep} onClick={()=>change('insertRowBefore',()=>focusRow(insertWordTableRow(table,top!,'before')))}>上方插入行</button><button disabled={busy||readOnly||!bottom||full} aria-label={`在${rowLabel}下方插入`} title="在下方插入行" onMouseDown={keep} onClick={()=>change('insertRowAfter',()=>focusRow(insertWordTableRow(table,bottom!,'after')))}>下方插入行</button><button disabled={busy||readOnly||!top} aria-label={`删除${rowLabel}`} title="删除当前行" onMouseDown={keep} onClick={()=>remote?change('deleteRow',()=>{}):removeRow()}>删除行</button></span>
   <span className="word-table-group" role="group" aria-label="列"><button disabled={busy||readOnly||!cell||columns>=WORD_TABLE_LIMITS.columns} onMouseDown={keep} onClick={()=>change('insertColumnBefore',()=>insertWordTableColumn(table,cell!,'before'))}>左侧插入列</button><button disabled={busy||readOnly||!cell||columns>=WORD_TABLE_LIMITS.columns} onMouseDown={keep} onClick={()=>change('insertColumnAfter',()=>insertWordTableColumn(table,cell!,'after'))}>右侧插入列</button><button disabled={busy||readOnly||!cell||columns<=1} onMouseDown={keep} onClick={()=>change('deleteColumn',()=>editing.focus(deleteWordTableColumn(table,cell!)))}>删除当前列</button></span>
   <span className="word-table-group" role="group" aria-label="合并"><button disabled={busy||readOnly||!span||span.column+span.colSpan>=columns} onMouseDown={keep} onClick={()=>change('mergeRight',()=>{mergeWordTableCell(table,cell!,'right');editing.focus(cell!);})}>向右合并</button><button disabled={busy||readOnly||!span||span.row+span.rowSpan>=rows.length} onMouseDown={keep} onClick={()=>change('mergeDown',()=>{mergeWordTableCell(table,cell!,'down');editing.focus(cell!);})}>向下合并</button><button disabled={busy||readOnly||!span||span.colSpan===1&&span.rowSpan===1} onMouseDown={keep} onClick={()=>change('split',()=>{splitWordTableCell(table,cell!);editing.focus(cell!);})}>拆分单元格</button></span>
   <span className="word-table-group" role="group" aria-label="表格"><button disabled={busy||readOnly||full} onMouseDown={keep} onClick={()=>change('appendRow',()=>focusRow(appendWordTableRow(table)))}>添加一行</button><button disabled={readOnly} onClick={onDelete}>删除表格</button></span>
   {!cell&&<span className="word-table-hint">点击单元格后可插入或删除行列</span>}{busy&&<span className="word-table-hint" role="status">正在提交表格修改…</span>}
  </header>
  {error&&<p role="alert">{error}{pending.current&&!busy&&<button type="button" onClick={()=>void send(pending.current!)}>重试表格操作</button>}</p>}{!rows.length?<p>表格暂无行。添加一行后可编辑和导出。</p>:<div className="word-table-scroll"><table><caption>第 {index+1} 个表格</caption><tbody>{list.map((row,r)=><tr key={key(row)}>{row.toArray().map((text,c)=>{const merge=layout.anchors.get(text);if(!merge||merge.row!==r||merge.column!==c)return null;return <td key={key(text)} colSpan={merge.colSpan>1?merge.colSpan:undefined} rowSpan={merge.rowSpan>1?merge.rowSpan:undefined}><WordTableCell text={text} label={`表格 ${index+1} 第 ${r+1} 行第 ${c+1} 列`} readOnly={readOnly} assertWritable={assertWritable} marker={markers.get(text)??''}/></td>;})}</tr>)}</tbody></table></div>}
 </section>;
}
