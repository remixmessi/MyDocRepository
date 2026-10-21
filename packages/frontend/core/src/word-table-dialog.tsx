import {useState} from 'react';
import {Y} from '@mydoc/crdt';
import {WORD_TABLE_LIMITS} from '../../../common/nbstore/src/office-word-table';
import {appendWordTable} from './office-word-table';
import {Modal} from './dialogs';

export function WordTableDialog({doc,readOnly,assertWritable,onClose}:{doc:Y.Doc;readOnly:boolean;assertWritable:()=>void;onClose:()=>void}){
 const [rows,setRows]=useState('2'),[columns,setColumns]=useState('2'),[error,setError]=useState('');
 const rowCount=Number(rows),columnCount=Number(columns);
 const valid=Number.isInteger(rowCount)&&rowCount>=1&&rowCount<=WORD_TABLE_LIMITS.rows&&Number.isInteger(columnCount)&&columnCount>=1&&columnCount<=WORD_TABLE_LIMITS.columns;
 return <Modal title="新建表格" onClose={onClose}><form onSubmit={event=>{
  event.preventDefault();if(readOnly||!valid)return;
  try{assertWritable();appendWordTable(doc,columnCount,rowCount);onClose();}catch(error){setError(error instanceof Error?error.message:'新建表格失败');}
 }}>
  <p>表格将添加到文档末尾，创建后可增删行；点击单元格可在左右插列或删除当前列。</p>
  <label className="form-field">行数<input autoFocus type="number" required min="1" max={WORD_TABLE_LIMITS.rows} step="1" value={rows} disabled={readOnly} onChange={event=>setRows(event.target.value)}/></label>
  <label className="form-field">列数<input type="number" required min="1" max={WORD_TABLE_LIMITS.columns} step="1" value={columns} disabled={readOnly} onChange={event=>setColumns(event.target.value)}/></label>
  <p role="status">{valid?`${rowCount} 行 × ${columnCount} 列，共 ${rowCount*columnCount} 个单元格`:'请输入 1–200 的整数行数和 1–50 的整数列数。'}</p>
  {readOnly&&<p role="status">当前文档只读。</p>}{error&&<p role="alert" className="form-error">{error}</p>}
  <div className="dialog-actions"><button type="button" onClick={onClose}>取消</button><button className="primary" disabled={readOnly||!valid}>插入表格</button></div>
 </form></Modal>;
}
