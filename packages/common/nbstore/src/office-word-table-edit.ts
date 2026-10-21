import {Y,LOCAL_ORIGIN} from '../../crdt/src/index';
import {createWordTable,isWordTable,WORD_TABLE_LIMITS,WORD_CELL_PARAGRAPH,wordTableRows,wordTableLayout,type WordBlock} from './office-word-table';

/**
 * Word table structure edits shared by the editor (personal workspaces, applied locally and undoable) and the server
 * (team workspaces, applied as semantic operations so concurrent row/column changes cannot leave a ragged table).
 */
function live(table:Y.Map<unknown>){if(!table.doc||!table.doc.getArray<WordBlock>('paragraphs').toArray().includes(table))throw new Error('表格已删除');return table.doc;}
function capacity(doc:Y.Doc,addedCells:number,newTable=false){
 const tables=doc.getArray<WordBlock>('paragraphs').toArray().filter(isWordTable);
 if(newTable&&tables.length>=WORD_TABLE_LIMITS.tables)throw new Error('每篇文档最多 100 个表格');
 const cells=tables.reduce((sum,table)=>sum+wordTableRows(table).length*Number(table.get('columns')),0);
 if(cells+addedCells>WORD_TABLE_LIMITS.cells)throw new Error('每篇文档最多 2 万个单元格');
}
export function appendWordTable(doc:Y.Doc,columns:number,rows:number,origin:unknown=LOCAL_ORIGIN){
 if(rows<1)throw new Error('新建表格至少需要一行');
 const table=createWordTable(columns,rows);capacity(doc,columns*rows,true);
 doc.transact(()=>doc.getArray<WordBlock>('paragraphs').push([table]),origin);return table;
}
function setSpan(cell:Y.Text,key:'colSpan'|'rowSpan',value:number){if(value>1)cell.setAttribute(key,value);else cell.removeAttribute(key);}
function insertRow(table:Y.Map<unknown>,index:number,origin:unknown){
 const doc=live(table),rows=wordTableRows(table),columns=Number(table.get('columns')),layout=wordTableLayout(table);
 if(rows.length>=WORD_TABLE_LIMITS.rows)throw new Error('表格最多 200 行');capacity(doc,columns);
 const row=new Y.Array<Y.Text>();row.push(Array.from({length:columns},()=>new Y.Text()));
 // A row inserted strictly inside a vertical merge joins it.
 doc.transact(()=>{for(const [cell,span]of layout.anchors)if(span.row<index&&span.row+span.rowSpan>index)setSpan(cell,'rowSpan',span.rowSpan+1);rows.insert(index,[row]);},origin);return row;
}
/** Removing the first row or column of a merge hands the content and the remaining span to the next covered cell. */
function shrinkSpans(table:Y.Map<unknown>,axis:'row'|'column',index:number){
 const rows=wordTableRows(table),layout=wordTableLayout(table);
 for(const [cell,span]of layout.anchors){
  const start=axis==='row'?span.row:span.column,size=axis==='row'?span.rowSpan:span.colSpan,key=axis==='row'?'rowSpan':'colSpan';
  if(size<2||index<start||index>=start+size)continue;
  if(index>start){setSpan(cell,key,size-1);continue;}
  const next=axis==='row'?rows.get(span.row+1).get(span.column):rows.get(span.row).get(span.column+1);
  next.applyDelta([...(next.length?[{delete:next.length}]:[]),...cell.toDelta()]);for(const [name,value]of Object.entries(cell.getAttributes()))next.setAttribute(name,value);setSpan(next,key,size-1);
 }
}
export function appendWordTableRow(table:Y.Map<unknown>,origin:unknown=LOCAL_ORIGIN){
 live(table);return insertRow(table,wordTableRows(table).length,origin);
}
export function insertWordTableRow(table:Y.Map<unknown>,anchor:Y.Array<Y.Text>,side:'before'|'after',origin:unknown=LOCAL_ORIGIN){
 live(table);if(side!=='before'&&side!=='after')throw new Error('插入位置无效');
 const index=wordTableRows(table).toArray().indexOf(anchor);if(index<0)throw new Error('目标行已删除，请重新选择');
 return insertRow(table,index+(side==='after'?1:0),origin);
}
export function deleteWordTableRow(table:Y.Map<unknown>,row:Y.Array<Y.Text>,origin:unknown=LOCAL_ORIGIN){
 const doc=live(table),rows=wordTableRows(table),index=rows.toArray().indexOf(row);if(index<0)throw new Error('行已删除');doc.transact(()=>{shrinkSpans(table,'row',index);rows.delete(index,1);},origin);
}
/** Resolve a column from a live cell rather than a stale screen index. */
export function wordTableColumn(table:Y.Map<unknown>,cell:Y.Text){
 live(table);for(const row of wordTableRows(table)){const column=row.toArray().indexOf(cell);if(column>=0)return column;}throw new Error('目标单元格已删除，请重新选择');
}
export function insertWordTableColumn(table:Y.Map<unknown>,cell:Y.Text,side:'before'|'after',origin:unknown=LOCAL_ORIGIN){
 const doc=live(table),rows=wordTableRows(table),columns=Number(table.get('columns')),layout=wordTableLayout(table);
 if(side!=='before'&&side!=='after')throw new Error('插入位置无效');
 // Relative to a merged cell, "after" means past its right edge.
 const span=layout.anchors.get(cell),index=wordTableColumn(table,cell)+(side==='after'?span?.colSpan??1:0);
 if(columns>=WORD_TABLE_LIMITS.columns)throw new Error('表格最多 50 列');capacity(doc,rows.length);
 doc.transact(()=>{for(const [anchor,span]of layout.anchors)if(span.column<index&&span.column+span.colSpan>index)setSpan(anchor,'colSpan',span.colSpan+1);for(const row of rows)row.insert(index,[new Y.Text()]);table.set('columns',columns+1);},origin);
}
export function deleteWordTableColumn(table:Y.Map<unknown>,cell:Y.Text,origin:unknown=LOCAL_ORIGIN){
 const doc=live(table),rows=wordTableRows(table),columns=Number(table.get('columns')),index=wordTableColumn(table,cell);
 if(columns<=1)throw new Error('至少保留一列，可删除整个表格');
 const next=rows.get(0).get(index===columns-1?index-1:index+1);
 doc.transact(()=>{shrinkSpans(table,'column',index);for(const row of rows)row.delete(index,1);table.set('columns',columns-1);},origin);return next;
}
/** Merge the next cell (or cells) to the right or below into this one; covered text joins as new cell paragraphs. */
export function mergeWordTableCell(table:Y.Map<unknown>,cell:Y.Text,direction:'right'|'down',origin:unknown=LOCAL_ORIGIN){
 const doc=live(table),rows=wordTableRows(table),layout=wordTableLayout(table),span=layout.anchors.get(cell);if(!span)throw new Error('目标单元格已删除或被合并，请重新选择');
 const from=direction==='right'?{row:span.row,column:span.column+span.colSpan}:{row:span.row+span.rowSpan,column:span.column},size=direction==='right'?span.rowSpan:span.colSpan;
 if(direction==='right'?from.column>=Number(table.get('columns')):from.row>=rows.length)throw new Error(direction==='right'?'右侧没有可合并的单元格':'下方没有可合并的单元格');
 const covered:Y.Text[]=[];for(let i=0;i<size;i++){const r=direction==='right'?from.row+i:from.row,c=direction==='right'?from.column:from.column+i,target=rows.get(r).get(c),other=layout.anchors.get(target);if(!other||other.colSpan>1||other.rowSpan>1)throw new Error('相邻区域已合并或不规则，请先拆分后再合并');covered.push(target);}
 if([cell,...covered].filter(text=>text.length).reduce((n,text)=>n+text.toString().split(WORD_CELL_PARAGRAPH).length,0)>WORD_TABLE_LIMITS.paragraphs)throw new Error('合并后单元格超过 100 段');
 doc.transact(()=>{
  for(const text of covered){if(text.length){const delta=text.toDelta();cell.applyDelta([{retain:cell.length},...(cell.length?[{insert:WORD_CELL_PARAGRAPH}]:[]),...delta]);text.delete(0,text.length);}for(const key of Object.keys(text.getAttributes()))text.removeAttribute(key);}
  setSpan(cell,direction==='right'?'colSpan':'rowSpan',(direction==='right'?span.colSpan:span.rowSpan)+1);
 },origin);
}
export function splitWordTableCell(table:Y.Map<unknown>,cell:Y.Text,origin:unknown=LOCAL_ORIGIN){
 const doc=live(table),span=wordTableLayout(table).anchors.get(cell);if(!span)throw new Error('目标单元格已删除或被合并，请重新选择');if(span.colSpan===1&&span.rowSpan===1)throw new Error('当前单元格未合并');
 doc.transact(()=>{cell.removeAttribute('colSpan');cell.removeAttribute('rowSpan');},origin);
}

/** Stable identity of a shared type: the Yjs id of the item that created it (never a screen index). */
export const wordItemIdentity=(value:{_item:Y.Item|null})=>{const id=value._item?.id;return id?`${id.client}-${id.clock}`:undefined;};
const IDENTITY=/^\d+-\d+$/;
/** Structure only: rows, cells and merges by identity. Cell text and paragraph formatting are not part of it. */
export function wordTableStructure(table:Y.Map<unknown>){
 return JSON.stringify([Number(table.get('columns')),wordTableRows(table).toArray().map(row=>[wordItemIdentity(row),row.toArray().map(cell=>[wordItemIdentity(cell),cell.getAttribute('colSpan')??1,cell.getAttribute('rowSpan')??1])])]);
}
/** Short, deterministic revision of a table's structure (FNV-1a 64 over the structure string; works in browser and server). */
export function wordTableRevision(table:Y.Map<unknown>){
 let a=0x811c9dc5,b=0x01000193;const text=wordTableStructure(table);
 for(let i=0;i<text.length;i++){const code=text.charCodeAt(i);a=Math.imul(a^code,0x01000193)>>>0;b=Math.imul(b^code^(a>>>8),0x5bd1e995)>>>0;}
 return a.toString(16).padStart(8,'0')+b.toString(16).padStart(8,'0')+text.length.toString(16);
}
export function findWordTable(doc:Y.Doc,identity:string){return doc.getArray<WordBlock>('paragraphs').toArray().find((block):block is Y.Map<unknown>=>isWordTable(block)&&wordItemIdentity(block)===identity);}
export function findWordTableCell(table:Y.Map<unknown>,identity:string){for(const row of wordTableRows(table))for(const cell of row)if(wordItemIdentity(cell)===identity)return {row,cell};}

export type WordTableAction='insertRowBefore'|'insertRowAfter'|'deleteRow'|'insertColumnBefore'|'insertColumnAfter'|'deleteColumn'|'mergeRight'|'mergeDown'|'split'|'appendRow';
export const WORD_TABLE_ACTIONS:WordTableAction[]=['insertRowBefore','insertRowAfter','deleteRow','insertColumnBefore','insertColumnAfter','deleteColumn','mergeRight','mergeDown','split','appendRow'];
/** `cell` is required except for appendRow; `base` is the table's structure revision the user acted on. */
export type WordTableOperation={operationId:string;table:string;base:string;action:WordTableAction;cell?:string};
export function validWordTableOperation(value:unknown):value is WordTableOperation{
 if(!value||typeof value!=='object'||Array.isArray(value))return false;const v=value as Record<string,unknown>,keys=Object.keys(v);
 return keys.every(key=>['operationId','table','base','action','cell'].includes(key))&&typeof v.operationId==='string'&&/^[a-zA-Z0-9-]{16,80}$/.test(v.operationId)&&typeof v.table==='string'&&IDENTITY.test(v.table)&&typeof v.base==='string'&&/^[0-9a-f]{16,40}$/.test(v.base)&&WORD_TABLE_ACTIONS.includes(v.action as WordTableAction)&&(v.action==='appendRow'?v.cell===undefined||typeof v.cell==='string'&&IDENTITY.test(v.cell):typeof v.cell==='string'&&IDENTITY.test(v.cell));
}
/**
 * Apply one semantic table operation and return the cell the caret should move to (by identity).
 * Rows are taken from the cell's merge area, so "above" uses its first row and "below" its last.
 */
export function applyWordTableOperation(doc:Y.Doc,input:WordTableOperation,origin:unknown){
 const table=findWordTable(doc,input.table);if(!table)throw new Error('表格已被其他成员删除');
 const rows=wordTableRows(table),layout=wordTableLayout(table),found=input.cell?findWordTableCell(table,input.cell):undefined;
 if(input.cell&&!found)throw new Error('目标单元格已删除，请重新选择');
 const cell=found?.cell,span=cell?layout.anchors.get(cell):undefined;if(cell&&!span)throw new Error('目标单元格已被合并，请重新选择');
 const column=span?.column??0,pick=(row:Y.Array<Y.Text>|undefined)=>row?wordItemIdentity(row.get(Math.min(column,row.length-1))):undefined;
 switch(input.action){
  case 'insertRowBefore':return pick(insertWordTableRow(table,rows.get(span!.row),'before',origin));
  case 'insertRowAfter':return pick(insertWordTableRow(table,rows.get(span!.row+span!.rowSpan-1),'after',origin));
  case 'appendRow':return pick(appendWordTableRow(table,origin));
  case 'deleteRow':{const at=span!.row;deleteWordTableRow(table,rows.get(at),origin);return pick(rows.length?rows.get(Math.min(at,rows.length-1)):undefined);}
  case 'insertColumnBefore':insertWordTableColumn(table,cell!,'before',origin);return input.cell;
  case 'insertColumnAfter':insertWordTableColumn(table,cell!,'after',origin);return input.cell;
  case 'deleteColumn':return wordItemIdentity(deleteWordTableColumn(table,cell!,origin));
  case 'mergeRight':mergeWordTableCell(table,cell!,'right',origin);return input.cell;
  case 'mergeDown':mergeWordTableCell(table,cell!,'down',origin);return input.cell;
  case 'split':splitWordTableCell(table,cell!,origin);return input.cell;
 }
}
