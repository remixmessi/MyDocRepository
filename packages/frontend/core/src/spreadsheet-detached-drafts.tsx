import {useState} from 'react';
import {Y} from '@mydoc/crdt';
import {workbookSheets} from '../../../common/nbstore/src/office-workbook';
import {cellDraftKey} from '../../../common/nbstore/src/spreadsheet-operation';
import {sheetColumnIds} from '../../../common/nbstore/src/sheet-columns';
import {downloadOfficeBytes} from './office-download';
import type {OfficeSpreadsheetDocument} from './workspace';
export function SpreadsheetDetachedDrafts({item}:{item:OfficeSpreadsheetDocument}){
 const [,refresh]=useState(0),sheets=workbookSheets(item.doc),detached=[...item.drafts].filter(([key])=>{try{const [encoded,column]=JSON.parse(key),position=Y.createAbsolutePositionFromRelativePosition(Y.decodeRelativePosition(Uint8Array.from(encoded)),item.doc),sheet=sheets.find(sheet=>sheet.sheet===position?.type);if(!position||!sheet||position.index>=sheet.sheet.length)return true;const c=sheetColumnIds(sheet.sheet).indexOf(column);return c<0||cellDraftKey(sheet.sheet,position.index,c)!==key;}catch{return true;}});
 if(!detached.length)return null;
 return <section className="sheet-detached-drafts" aria-label="已删除单元格的草稿"><p role="alert">有 {detached.length} 份草稿对应的行列已删除。输入仍保留，可下载后复制到其他单元格，或明确放弃。</p><button onClick={()=>downloadOfficeBytes('未提交单元格草稿.json',new TextEncoder().encode(JSON.stringify({type:'mydoc-sheet-detached-drafts',version:1,resource:item.id,drafts:detached.map(([,draft])=>draft)},null,2)),'application/json')}>下载已删除单元格草稿</button><ul>{detached.map(([key,draft])=><li key={key}><pre>{draft.value}</pre><button onClick={()=>{item.drafts.delete(key);refresh(v=>v+1);}}>放弃此单元格草稿</button></li>)}</ul></section>;
}
