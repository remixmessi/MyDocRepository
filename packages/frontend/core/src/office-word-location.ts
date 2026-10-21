import {Y} from '@mydoc/crdt';
import {isWordTable,wordTableRows,type WordBlock} from '../../../common/nbstore/src/office-word-table';

/** Resolve membership through the current object graph, never a saved row index. */
export function assertWordTextLive(text:Y.Text){
 const doc=text.doc;
 if(doc){
  const blocks=doc.getArray<WordBlock>('paragraphs');
  if(text.parent===blocks&&blocks.toArray().includes(text))return;
  const row=text.parent,rows=row?.parent,table=rows?.parent;
  if(row instanceof Y.Array&&rows instanceof Y.Array&&isWordTable(table)&&table.parent===blocks&&blocks.toArray().includes(table)&&wordTableRows(table)===rows&&rows.toArray().includes(row)&&row.toArray().includes(text))return;
 }
 throw new Error('段落或单元格已删除，请重新选择');
}
