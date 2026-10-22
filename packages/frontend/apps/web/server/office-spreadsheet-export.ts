import {workbookFormulaSnapshot,validWorksheetName} from '../../../../common/nbstore/src/office-workbook';
import {calculateFormulaWorkbookTyped} from '../../../../common/nbstore/src/spreadsheet-formula';
import {spreadsheetCellType} from '../../../../common/nbstore/src/spreadsheet-cell-type';
import * as Y from 'yjs';
import {validOfficeComment,type OfficeCommentThread} from '../../../../common/nbstore/src/office-comments';
import {sheetColumnIds,sheetRowIdentity} from '../../../../common/nbstore/src/sheet-columns';
import {strToU8,zipSync} from 'fflate';
import {validateSpreadsheetUpdate} from './office-spreadsheet';
import {toXlsxFormula} from '../../../../common/nbstore/src/spreadsheet-functions';
import {readOfficePackage,readOfficeXml} from './office-package';
import {conditionalXml,dxfRegistry,sheetCharts} from './office-xlsx-objects';
import {previewFormulaWorkbook} from '../../../../common/nbstore/src/spreadsheet-formula';
import {sheetCellFormat,resolveSheetMerges,formatSheetValue,isDateNumberFormat,type SheetCellFormat} from '../../../../common/nbstore/src/spreadsheet-format';
/** Collect distinct cell formats into a minimal styles part; style 0 stays the unformatted default. */
function styleRegistry(){
 const fonts=['<font><sz val="11"/><name val="Calibri"/></font>'],fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'],formats:string[]=[],xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'],index=new Map<string,number>();
 const add=(list:string[],xml:string)=>{const found=list.indexOf(xml);if(found>=0)return found;list.push(xml);return list.length-1;};
 return {
  style(format:SheetCellFormat|undefined){
   if(!format)return 0;const key=JSON.stringify(format),known=index.get(key);if(known!==undefined)return known;
   const font=format.bold||format.italic||format.underline||format.color?add(fonts,`<font>${format.bold?'<b/>':''}${format.italic?'<i/>':''}${format.underline?'<u/>':''}<sz val="11"/>${format.color?`<color rgb="FF${format.color}"/>`:''}<name val="Calibri"/></font>`):0;
   const fill=format.fill?add(fills,`<fill><patternFill patternType="solid"><fgColor rgb="FF${format.fill}"/><bgColor indexed="64"/></patternFill></fill>`):0;
   const numFmt=format.numberFormat==='@'?49:format.numberFormat?164+add(formats,format.numberFormat):0;
   const xf=`<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="0" xfId="0"${numFmt?' applyNumberFormat="1"':''}${font?' applyFont="1"':''}${fill?' applyFill="1"':''}${format.align?` applyAlignment="1"><alignment horizontal="${format.align}"/></xf>`:'/>'}`;
   if(xfs.length>=64000)throw new Error('单元格样式组合超过 XLSX 上限');xfs.push(xf);index.set(key,xfs.length-1);return xfs.length-1;
  },
  xml(dxfs=''){return '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+(formats.length?`<numFmts count="${formats.length}">${formats.map((code,i)=>`<numFmt numFmtId="${164+i}" formatCode="${xmlText(code)}"/>`).join('')}</numFmts>`:'')+`<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${dxfs}</styleSheet>`;},
 };
}

// SpreadsheetML escapes would otherwise turn a literal _x000A_ into a newline.
function xmlText(value:string){
 for(const char of value){const n=char.codePointAt(0)!;if(!(n===9||n===10||n===13||n>=32&&n<=0xd7ff||n>=0xe000&&n<=0xfffd||n>=0x10000&&n<=0x10ffff))throw new Error('单元格包含无法导出为 XLSX 的字符');}
 return value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll('\r','&#13;');
}
function text(value:string){return xmlText(value.replace(/_(?=x[0-9a-f]{4}_)/gi,'_x005F_'));}
/**
 * Cell discussions become Excel notes (comments.xml plus the VML shape Excel needs to show them). A thread's messages
 * are joined into one note — "作者：内容" per line, with a "[已解决]" prefix — because notes have no reply structure.
 */
function sheetNotes(doc:Y.Doc,sheet:{id:string;sheet:Y.Array<Y.Array<string>>}){
 if(!doc.share.has('officeComments'))return [];
 const rows=new Map(sheet.sheet.toArray().map((row,r)=>[sheetRowIdentity(row),r])),columns=sheetColumnIds(sheet.sheet),notes=new Map<string,{row:number;column:number;authors:string[];lines:string[]}>();
 const threads=[...doc.getMap<OfficeCommentThread>('officeComments').values()].filter(validOfficeComment).sort((a,b)=>a.messages[0].created-b.messages[0].created||a.id.localeCompare(b.id));
 for(const thread of threads){const anchor=thread.anchor;if(anchor?.kind!=='cell'||anchor.sheet!==sheet.id)continue;const row=rows.get(anchor.row),columnIndex=columns.indexOf(anchor.column);if(row===undefined||columnIndex<0)continue;
  const key=row+':'+columnIndex,note=notes.get(key)??{row,column:columnIndex,authors:[],lines:[]};if(note.lines.length)note.lines.push('');
  thread.messages.forEach((message,i)=>{note.lines.push((i===0&&thread.resolved?'[已解决] ':'')+message.authorName+'：'+message.body);if(!note.authors.includes(message.authorName))note.authors.push(message.authorName);});notes.set(key,note);}
 return [...notes.values()].sort((a,b)=>a.row-b.row||a.column-b.column);
}
function column(index:number){let name='';for(let n=index+1;n;n=Math.floor((n-1)/26))name=String.fromCharCode(65+(n-1)%26)+name;return name;}
/** A declared text-only single-sheet copy. No formula execution or fabricated caches. */
export function exportXlsxText(snapshot:Uint8Array){return exportXlsx(snapshot,'text');}
export function exportXlsxFormula(snapshot:Uint8Array,cache=false){return exportXlsx(snapshot,'formula',cache);}
function exportXlsx(snapshot:Uint8Array,mode:'text'|'formula',cache=false){
 validateSpreadsheetUpdate(null,snapshot);const doc=new Y.Doc();
 try{
  Y.applyUpdate(doc,snapshot);if(doc.store.pendingStructs||doc.store.pendingDs)throw new Error('表格状态缺少依赖');
  const {sheets,book}=workbookFormulaSnapshot(doc),byName=new Map(sheets.map(sheet=>[sheet.name,sheet]));
  if(sheets.some(sheet=>!validWorksheetName(sheet.name)))throw new Error('工作表名称不符合 XLSX 限制，无法保持名称导出；可选择 CSV/TSV 或办公文件备份');
  const computed=mode==='formula'?calculateFormulaWorkbookTyped(book,(name,r,c)=>spreadsheetCellType(doc,r,c,byName.get(name)!.id)):undefined;
  let stringSize=0,count=0;const strings:string[]=[],worksheetFiles:Record<string,Uint8Array>={},noteFiles:Record<string,Uint8Array>={},noteTypes:string[]=[];const indexes=new Map<string,number>(),styles=styleRegistry(),dxfs=dxfRegistry(),objectFiles:Record<string,Uint8Array>={},objectTypes:string[]=[];let chartNumber=1,drawingNumber=1;
  // Chart caches and conditional formats use the values a viewer sees (formula results), in both copy modes.
  const shownValues=previewFormulaWorkbook(book,(name,r,c)=>spreadsheetCellType(doc,r,c,byName.get(name)!.id));
  const stringIndex=(value:string)=>{count++;const previous=indexes.get(value);if(previous!==undefined)return previous;const xml='<si><t xml:space="preserve">'+text(value)+'</t></si>';stringSize+=strToU8(xml).length;if(stringSize>8*1024*1024-512)throw new Error('共享文本 XML 超过 8 MB');const index=strings.length;indexes.set(value,index);strings.push(xml);return index;};
  for(const [sheetIndex,selected]of sheets.entries()){
   const rows=selected.sheet.toArray(),sheetName=selected.name,calculated=computed?.[sheetName],parts:string[]=[],merges=resolveSheetMerges(doc,selected.id,selected.sheet);let size=0;
   if(rows.some(row=>row.length>16384))throw new Error('XLSX 列数超过 16384');
  const cellXml=(value:string,r:number,c:number)=>{
   const ref=column(c)+(r+1),kind=spreadsheetCellType(doc,r,c,selected.id),format=sheetCellFormat(doc,selected.id,selected.sheet,r,c),style=styles.style(format),s=style?` s="${style}"`:'';
   // The text copy stores what the cell shows, so dates are not reduced to serial numbers.
   if(mode==='text'&&isDateNumberFormat(format?.numberFormat)&&kind!=='text')value=formatSheetValue(value,format!.numberFormat);
   if(mode==='formula'&&kind!=='text'){
    if(value.trimStart().startsWith('=')){
     const result=calculated?.[r][c]??0,type=typeof result==='string'?' t="str"':typeof result==='boolean'?' t="b"':'';
     const cached=typeof result==='boolean'?result?'1':'0':xmlText(String(result));
     return `<c r="${ref}"${s}${cache?type:''}><f>${xmlText(toXlsxFormula(value.trimStart().slice(1)))}</f>${cache?'<v>'+cached+'</v>':''}</c>`;
    }
    const number=value.trim();
    if(!number)return `<c r="${ref}"${s}/>`;
    if(/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(number)){
     const significant=number.split(/[eE]/)[0].replace(/^[+-]/,'').replace('.','').replace(/^0+/,'').replace(/0+$/,'');
     if(!Number.isFinite(Number(number))||significant.length>15||Number(number)===0&&significant)throw new Error(`${ref} 自动类型数值超出支持的精度，请改为文本或导出 XLSX 文本副本`);
     return `<c r="${ref}"${s}><v>${xmlText(number)}</v></c>`;
    }
   }
   if(!value&&style)return `<c r="${ref}"${s}/>`;return `<c r="${ref}"${s} t="s"><v>${stringIndex(value)}</v></c>`;
  };
  for(const [r,row]of rows.entries()){
   const xml=`<row r="${r+1}">${row.toArray().map((value,c)=>cellXml(value,r,c)).join('')}</row>`;
   size+=strToU8(xml).length;if(size>8*1024*1024-512)throw new Error('工作表 XML 超过 8 MB');parts.push(xml);
  }
   const notes=sheetNotes(doc,selected),n=sheetIndex+1,sheetRelations:string[]=[];
   const shown=(r:number,c:number)=>shownValues[sheetName]?.[r]?.[c]??'',label=(r:number,c:number)=>formatSheetValue(shown(r,c),sheetCellFormat(doc,selected.id,selected.sheet,r,c)?.numberFormat);
   const conditional=conditionalXml(doc,selected.id,selected.sheet,style=>dxfs.id(style)),charts=sheetCharts(doc,selected.id,sheetName,selected.sheet,drawingNumber,chartNumber,shown,label);
   if(charts){drawingNumber++;chartNumber=charts.next;for(const [name,xml]of Object.entries(charts.files))objectFiles[name]=strToU8(xml);objectTypes.push(...charts.types);sheetRelations.push(charts.relation);}
   if(notes.length){
    const authors=[...new Set(notes.flatMap(note=>note.authors))];
    const commentsXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>'+authors.map(name=>'<author>'+text(name)+'</author>').join('')+'</authors><commentList>'+notes.map(note=>`<comment ref="${column(note.column)+(note.row+1)}" authorId="${authors.indexOf(note.authors[0])}"><text>${note.lines.map((line,i)=>`<r><t xml:space="preserve">${text((i?'\n':'')+line)}</t></r>`).join('')}</text></comment>`).join('')+'</commentList></comments>';
    if(strToU8(commentsXml).length>8*1024*1024)throw new Error('导出批注 XML 超过 8 MB');
    const shape=(note:{row:number;column:number},i:number)=>`<v:shape id="_x0000_s${1025+i}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:160pt;height:70pt;z-index:${i+1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>${note.row}</x:Row><x:Column>${note.column}</x:Column></x:ClientData></v:shape>`;
    noteFiles[`xl/comments${n}.xml`]=strToU8(commentsXml);
    noteFiles[`xl/drawings/vmlDrawing${n}.vml`]=strToU8(`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${n}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${notes.map(shape).join('')}</xml>`);
    sheetRelations.push(`<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${n}.xml"/>`,`<Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${n}.vml"/>`);
    noteTypes.push(`<Override PartName="/xl/comments${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`);
   }
   if(sheetRelations.length)noteFiles[`xl/worksheets/_rels/sheet${n}.xml.rels`]=strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRelations.join('')}</Relationships>`);
   // CT_Worksheet order: sheetData, mergeCells, conditionalFormatting, …, drawing, legacyDrawing.
   worksheetFiles[`xl/worksheets/sheet${n}.xml`]=strToU8('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>'+parts.join('')+'</sheetData>'+(merges.anchors.size?`<mergeCells count="${merges.anchors.size}">${[...merges.anchors.values()].map(m=>`<mergeCell ref="${column(m.column)+(m.row+1)}:${column(m.column+m.colSpan-1)+(m.row+m.rowSpan)}"/>`).join('')}</mergeCells>`:'')+conditional+(charts?'<drawing r:id="rIdDrawing"/>':'')+(notes.length?'<legacyDrawing r:id="rIdVml"/>':'')+'</worksheet>');
  }
  const files={
   '[Content_Types].xml':strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+sheets.map((_,i)=>'<Override PartName="/xl/worksheets/sheet'+(i+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('')+(noteTypes.length?'<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>':'')+noteTypes.join('')+objectTypes.join('')+'</Types>'),
   '_rels/.rels':strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
   'xl/workbook.xml':strToU8('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sheets.map((sheet,i)=>'<sheet name="'+xmlText(sheet.name)+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>').join('')+'</sheets>'+ (mode==='formula'?'<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/>':'')+'</workbook>'),
   'xl/_rels/workbook.xml.rels':strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+sheets.map((_,i)=>'<Relationship Id="rId'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(i+1)+'.xml"/>').join('')+'<Relationship Id="rIdStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
   'xl/sharedStrings.xml':strToU8('<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+count+'" uniqueCount="'+strings.length+'">'+strings.join('')+'</sst>'),
   'xl/styles.xml':strToU8(styles.xml(dxfs.xml())),
   ...worksheetFiles,
   ...noteFiles,
   ...objectFiles,
  };
  const bytes=zipSync(files);const verified=readOfficePackage(bytes);for(const [name,value]of Object.entries(verified))if(/\.(xml|rels)$/.test(name))readOfficeXml(value);return bytes;
 }finally{doc.destroy();}
}
