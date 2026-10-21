import {WORD_TABLE_LIMITS} from '../../../../common/nbstore/src/office-word-table';
/** One grid position. `fill` marks columns covered by a gridSpan to the left; `continued` marks a vMerge continuation. */
export type WordTableCellSource={paragraphs:string[];colSpan:number;rowSpan:number;fill:boolean;continued:boolean};
function cellProperties(xml:string){return /<w:tcPr(?:\s[^>]*)?>([\s\S]*?)<\/w:tcPr>/.exec(xml)?.[1]??'';}
function value(properties:string,tag:string){const element=new RegExp(`<w:${tag}(?=\\s|/|>)([^>]*?)(?:/>|>)`).exec(properties);if(!element)return;return /\bw:val\s*=\s*(["'])(.*?)\1/.exec(element[1])?.[2]??'';}
/** Expand OOXML gridSpan / vMerge into a rectangular grid of anchors and covered cells. */
export function readWordTable(xml:string){
 if(/<w:tbl(?=\s|\/|>)/.test(xml))throw new Error('暂不支持嵌套 Word 表格');
 if(/<w:hMerge(?=\s|\/|>)/.test(xml))throw new Error('暂不支持旧式横向合并 Word 单元格，请改用跨列合并');
 const rows:WordTableCellSource[][]=[];
 for(const row of xml.matchAll(/<w:tr(?:\s[^>]*?)?(?:\s*\/>|>([\s\S]*?)<\/w:tr>)/g)){
  if(rows.length>=WORD_TABLE_LIMITS.rows)throw new Error('Word 表格超过 200 行');
  if(/<w:(?:gridBefore|gridAfter)(?=\s|\/|>)/.test(row[1]??''))throw new Error('暂不支持 Word 表格省略单元格');
  const cells:WordTableCellSource[]=[];
  for(const cell of (row[1]??'').matchAll(/<w:tc(?:\s[^>]*?)?(?:\s*\/>|>([\s\S]*?)<\/w:tc>)/g)){
   const content=cell[1]??'',properties=cellProperties(content),span=value(properties,'gridSpan'),merge=value(properties,'vMerge');
   const colSpan=span===undefined?1:Number(span);if(!Number.isInteger(colSpan)||colSpan<1||colSpan>WORD_TABLE_LIMITS.columns)throw new Error('Word 跨列合并数值无效');
   if(merge!==undefined&&!['','continue','restart'].includes(merge))throw new Error('Word 跨行合并数值无效');
   const paragraphs=[...content.matchAll(/<w:p(?:\s[^>]*?)?(?:\s*\/>|>([\s\S]*?)<\/w:p>)/g)].map(p=>p[1]??'');
   if(!paragraphs.length)throw new Error('Word 单元格缺少段落');
   const continued=merge===''||merge==='continue';
   // Continuation cells are hidden in Word; refuse rather than silently drop any text they still hold.
   if(continued&&paragraphs.some(p=>/<w:t(?:\s[^>]*)?>[^<]+<\/w:t>|<w:drawing\b/.test(p)))throw new Error('Word 跨行合并的被合并单元格含有内容，请先在 Word 中清理');
   cells.push({paragraphs:continued?['']:paragraphs,colSpan,rowSpan:1,fill:false,continued});
   for(let i=1;i<colSpan;i++)cells.push({paragraphs:[''],colSpan:1,rowSpan:1,fill:true,continued});
   if(cells.length>WORD_TABLE_LIMITS.columns)throw new Error('Word 表格超过 50 列');
  }
  if(!cells.length||rows.length&&cells.length!==rows[0].length)throw new Error('Word 表格必须是完整矩形');rows.push(cells);
 }
 if(!rows.length)throw new Error('Word 表格缺少行');
 const grid=[...xml.matchAll(/<w:gridCol(?=\s|\/|>)/g)].length;if(grid&&grid!==rows[0].length)throw new Error('Word 表格网格与列数不一致');
 // Each continuation extends the anchor that starts the vertical run at the same grid column with the same width.
 rows.forEach((row,r)=>row.forEach((cell,c)=>{
  if(cell.fill||!cell.continued)return;
  let above=r-1;while(above>=0&&rows[above][c].continued&&!rows[above][c].fill)above--;
  const anchor=rows[above]?.[c];if(!anchor||anchor.fill||anchor.continued||anchor.colSpan!==cell.colSpan)throw new Error('Word 跨行合并缺少起始单元格或跨列宽度不一致');anchor.rowSpan++;
 }));
 return rows;
}
