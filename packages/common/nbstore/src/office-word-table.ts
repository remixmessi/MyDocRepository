// Also loaded by Vite's bundled server configuration; keep the workspace source
// in that bundle rather than externalizing a package whose entry is TypeScript.
import {isWordImage,validateWordImage} from './office-image';
import {Y} from '../../crdt/src/index';
export type WordBlock=Y.Text|Y.Map<unknown>;
export const WORD_TABLE_LIMITS={tables:100,rows:200,columns:50,cells:20000,paragraphs:100};
export function isWordTable(value:unknown):value is Y.Map<unknown>{return value instanceof Y.Map&&value.get('kind')==='table';}
export function createWordTable(columns:number,rows=2){
 if(!Number.isInteger(columns)||columns<1||columns>WORD_TABLE_LIMITS.columns||!Number.isInteger(rows)||rows<0||rows>WORD_TABLE_LIMITS.rows)throw new Error('Word 表格行列数无效');
 const table=new Y.Map<unknown>(),data=new Y.Array<Y.Array<Y.Text>>();table.set('kind','table');table.set('columns',columns);table.set('rows',data);
 for(let i=0;i<rows;i++){const row=new Y.Array<Y.Text>();row.push(Array.from({length:columns},()=>new Y.Text()));data.push([row]);}return table;
}
export function wordTableRows(table:Y.Map<unknown>){return table.get('rows') as Y.Array<Y.Array<Y.Text>>;}
/** Paragraph separator inside one table cell; `\n` remains a soft line break. */
export const WORD_CELL_PARAGRAPH='\u2029';
export function wordCellSpan(cell:Y.Text){const value=(key:'colSpan'|'rowSpan')=>{const n=cell.getAttribute(key);return Number.isInteger(n)&&Number(n)>1?Number(n):1;};return {colSpan:value('colSpan'),rowSpan:value('rowSpan')};}
export type WordTableLayout={anchors:Map<Y.Text,{row:number;column:number;colSpan:number;rowSpan:number}>;owner:(Y.Text|undefined)[][]};
/**
 * Resolve merges row-major. A span only applies when every cell it covers is empty and has no span of its own,
 * so concurrent or stale merges degrade to separate cells and never hide text.
 */
export function wordTableLayout(table:Y.Map<unknown>):WordTableLayout{
 const rows=wordTableRows(table).toArray().map(row=>row.toArray()),owner=rows.map(row=>row.map(()=>undefined as Y.Text|undefined)),anchors:WordTableLayout['anchors']=new Map();
 rows.forEach((row,r)=>row.forEach((cell,c)=>{
  if(owner[r][c])return;let {colSpan,rowSpan}=wordCellSpan(cell);colSpan=Math.min(colSpan,row.length-c);rowSpan=Math.min(rowSpan,rows.length-r);
  const covered=(rr:number,cc:number)=>(rr!==r||cc!==c)&&(owner[rr]?.[cc]!==undefined||rows[rr][cc].length>0||wordCellSpan(rows[rr][cc]).colSpan>1||wordCellSpan(rows[rr][cc]).rowSpan>1);
  for(let rr=r;rr<r+rowSpan&&(colSpan>1||rowSpan>1);rr++)for(let cc=c;cc<c+colSpan;cc++)if(covered(rr,cc)){colSpan=1;rowSpan=1;}
  for(let rr=r;rr<r+rowSpan;rr++)for(let cc=c;cc<c+colSpan;cc++)owner[rr][cc]=cell;anchors.set(cell,{row:r,column:c,colSpan,rowSpan});
 }));
 return {anchors,owner};
}
export function validateWordTable(table:Y.Map<unknown>){
 const columns=table.get('columns'),rows=table.get('rows');
 if(table._start||table.size!==3||table.get('kind')!=='table'||!Number.isInteger(columns)||Number(columns)<1||Number(columns)>WORD_TABLE_LIMITS.columns||!(rows instanceof Y.Array)||rows._map.size||rows.length>WORD_TABLE_LIMITS.rows)throw new Error('Word 表格结构无效');
 const cells:Y.Text[]=[];for(const row of rows.toArray()){if(!(row instanceof Y.Array)||row._map.size||row.length!==columns)throw new Error('Word 表格行列结构无效');for(const cell of row.toArray()){if(!(cell instanceof Y.Text))throw new Error('Word 单元格必须是文字');cells.push(cell);}}
 return cells;
}
export function wordBlockTexts(blocks:WordBlock[]){return blocks.flatMap(block=>block instanceof Y.Text?[block]:isWordImage(block)?(validateWordImage(block,true),[]):validateWordTable(block));}
