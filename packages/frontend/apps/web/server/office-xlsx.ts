import {FIRST_WORKSHEET,MAX_WORKBOOK_SHEETS,validWorksheetName} from '../../../../common/nbstore/src/office-workbook';
import * as Y from 'yjs';
import {posix} from 'node:path';
import {readOfficePackage} from './office-package';
import {xlsxXml,attribute,children,child,stringValue,type XlsxXml} from './office-xlsx-xml';
import {setSpreadsheetCellType,type SpreadsheetCellType} from '../../../../common/nbstore/src/spreadsheet-cell-type';
import {validXlsxSource,type XlsxSource} from '../../../../common/nbstore/src/office-xlsx-source';
import {sheetCellKey,validSheetCellFormat,validSheetNumberFormat,validSheetFontSize,isDateNumberFormat,SHEET_DEFAULT_FONT_SIZE,SHEET_FORMAT_LIMITS,type SheetCellFormat} from '../../../../common/nbstore/src/spreadsheet-format';
import {sheetColumnIds,sheetRowIdentity} from '../../../../common/nbstore/src/sheet-columns';
import {fromXlsxFormula} from '../../../../common/nbstore/src/spreadsheet-functions';
import {readDxfs,readSheetObjects} from './office-xlsx-objects';
import {sheetDimensionKey,excelWidthToPixels,pointsToPixels,SHEET_DIMENSION_LIMITS} from '../../../../common/nbstore/src/spreadsheet-dimensions';
import {sheetRangeKeys,validSheetObject,SHEET_OBJECT_LIMITS,type SheetObject} from '../../../../common/nbstore/src/spreadsheet-objects';
// Built-in number formats with locale-neutral codes (ECMA-376 18.8.30); others fall back to General.
const BUILTIN_FORMATS:Record<number,string>={1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',14:'yyyy-mm-dd',15:'d-mmm-yy',16:'d-mmm',17:'mmm-yy',18:'h:mm AM/PM',19:'h:mm:ss AM/PM',20:'h:mm',21:'h:mm:ss',22:'yyyy-mm-dd h:mm',45:'mm:ss',46:'[h]:mm:ss',47:'mm:ss.0',49:'@'};
/** Read cellXfs into the subset the editor supports; theme/indexed colours and font families are not carried over. */
function readStyles(bytes:Uint8Array|undefined){
 if(!bytes)return [] as (SheetCellFormat|undefined)[];const root=xlsxXml(bytes);if(root.name!=='styleSheet')throw new Error('XLSX 样式结构无效');
 const rgb=(node:XlsxXml|undefined)=>{const value=node&&attribute(node,'rgb');return value&&/^[0-9A-Fa-f]{8}$/.test(value)?value.slice(2).toUpperCase():undefined;};
 const codes=new Map<number,string>();for(const format of child(root,'numFmts')?children(child(root,'numFmts')!,'numFmt'):[]){const id=Number(attribute(format,'numFmtId')),code=attribute(format,'formatCode');if(Number.isInteger(id)&&code&&validSheetNumberFormat(code))codes.set(id,code);}
 const fonts=(child(root,'fonts')?children(child(root,'fonts')!,'font'):[]).map(font=>{const on=(name:string)=>{const node=child(font,name);return !!node&&!['0','false'].includes(attribute(node,'val')??'1');};const u=child(font,'u'),size=Number(attribute(child(font,'sz')??font,'val'));return {bold:on('b'),italic:on('i'),strike:on('strike'),underline:!!u&&(attribute(u,'val')??'single')==='single',color:rgb(child(font,'color')),fontSize:Number.isFinite(size)&&size!==SHEET_DEFAULT_FONT_SIZE?Math.round(size*10)/10:undefined};});
 // Border sides keep the styles the editor draws (other dash patterns map to the nearest one); theme colours become black.
 const lines:Record<string,string>={thin:'thin',hair:'thin',medium:'medium',thick:'thick',dashed:'dashed',mediumDashed:'dashed',dashDot:'dashed',mediumDashDot:'dashed',dashDotDot:'dotted',mediumDashDotDot:'dotted',slantDashDot:'dashed',dotted:'dotted',double:'double'};
 const borders=(child(root,'borders')?children(child(root,'borders')!,'border'):[]).map(border=>Object.fromEntries((['left','right','top','bottom'] as const).flatMap(name=>{const node=child(border,name),style=node&&lines[attribute(node,'style')??''];if(!style)return [];const hex=rgb(node.children.find(item=>item.name==='color'));return [['border'+name[0].toUpperCase()+name.slice(1),style+(hex&&hex!=='000000'?':'+hex:'')]];})));
 const fills=(child(root,'fills')?children(child(root,'fills')!,'fill'):[]).map(fill=>{const pattern=child(fill,'patternFill');return pattern&&attribute(pattern,'patternType')==='solid'?rgb(child(pattern,'fgColor')):undefined;});
 const xfs=child(root,'cellXfs')?children(child(root,'cellXfs')!,'xf'):[];if(xfs.length>64000)throw new Error('XLSX 单元格样式过多');
 return xfs.map(xf=>{
  const font=fonts[Number(attribute(xf,'fontId')??0)],fill=fills[Number(attribute(xf,'fillId')??0)],border=borders[Number(attribute(xf,'borderId')??0)],id=Number(attribute(xf,'numFmtId')??0),alignment=child(xf,'alignment'),align=alignment&&attribute(alignment,'horizontal'),vertical=alignment&&attribute(alignment,'vertical'),verticalAlign=vertical==='center'?'middle':vertical==='top'?'top':undefined;
  const format:Record<string,unknown>={...(font?.bold?{bold:true}:{}),...(font?.italic?{italic:true}:{}),...(font?.underline?{underline:true}:{}),...(font?.strike?{strike:true}:{}),...(font?.fontSize&&validSheetFontSize(font.fontSize)?{fontSize:font.fontSize}:{}),...(font?.color&&font.color!=='000000'?{color:font.color}:{}),...(fill?{fill}:{}),...(['left','center','right'].includes(align??'')?{align}:{}),...(verticalAlign?{verticalAlign}:{}),...(alignment&&['1','true'].includes(attribute(alignment,'wrapText')??'')?{wrap:true}:{}),...border};
  const code=codes.get(id)??BUILTIN_FORMATS[id];if(id&&code)format.numberFormat=code;
  return validSheetCellFormat(format)?format as SheetCellFormat:undefined;
 });
}
function cellRange(ref:string){const match=/^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/.exec(ref);if(!match)return;const col=(letters:string)=>[...letters].reduce((n,char)=>n*26+char.charCodeAt(0)-64,0)-1;const r1=Number(match[2])-1,c1=col(match[1]),r2=match[4]?Number(match[4])-1:r1,c2=match[3]?col(match[3]):c1;return {row:Math.min(r1,r2),column:Math.min(c1,c2),rowSpan:Math.abs(r2-r1)+1,colSpan:Math.abs(c2-c1)+1};}
function targetPath(target:string|undefined){if(!target||/[\\:#?]/.test(target))throw new Error('XLSX 工作表关系路径无效');const path=target.startsWith('/')?target.slice(1):posix.normalize(posix.join('xl',target));if(!path.startsWith('xl/')||path.split('/').includes('..'))throw new Error('XLSX 工作表关系路径无效');return path;}
function selectedSheets(files:Record<string,Uint8Array>){
 if(!files['xl/workbook.xml']){const paths=Object.keys(files).filter(name=>/^xl\/worksheets\/sheet\d+\.xml$/.test(name)).sort((a,b)=>Number(/sheet(\d+)/.exec(a)![1])-Number(/sheet(\d+)/.exec(b)![1]));if(!paths.length)throw new Error('XLSX 缺少工作表');return {sheets:paths.map((path,i)=>({path,name:'工作表 '+(i+1)})),shared:files['xl/sharedStrings.xml']};}
 const workbook=xlsxXml(files['xl/workbook.xml']);if(workbook.name!=='workbook')throw new Error('XLSX 工作簿结构无效');const sheets=child(workbook,'sheets'),list=sheets?children(sheets,'sheet'):[];if(!list.length||list.length>MAX_WORKBOOK_SHEETS)throw new Error('XLSX 工作表数量无效');
 const relations=files['xl/_rels/workbook.xml.rels'];if(!relations)throw new Error('XLSX 缺少工作簿关系');const root=xlsxXml(relations);if(root.name!=='Relationships')throw new Error('XLSX 工作簿关系无效');const map=new Map<string,XlsxXml>();for(const relation of children(root,'Relationship')){const id=attribute(relation,'Id');if(!id||map.has(id))throw new Error('XLSX 工作簿关系重复或无效');map.set(id,relation);}
 const sheetsSelected=list.map(sheet=>{const relation=map.get(attribute(sheet,'id')??'');if(!relation||attribute(relation,'TargetMode')==='External'||!attribute(relation,'Type')?.endsWith('/worksheet'))throw new Error('XLSX 工作表关系不受支持');return {path:targetPath(attribute(relation,'Target')),name:attribute(sheet,'name')??''};});
 const strings=[...map.values()].filter(value=>attribute(value,'Type')?.endsWith('/sharedStrings'));if(strings.length>1)throw new Error('XLSX 共享字符串关系重复');if(strings[0]&&attribute(strings[0],'TargetMode')==='External')throw new Error('XLSX 不支持外部共享字符串');const sharedPath=strings[0]?targetPath(attribute(strings[0],'Target')):'xl/sharedStrings.xml';
 return {sheets:sheetsSelected,shared:files[sharedPath]};
}
function parseSheet(bytes:Uint8Array,strings:string[],styles:(SheetCellFormat|undefined)[],date1904:boolean){
 const worksheet=xlsxXml(bytes);if(worksheet.name!=='worksheet')throw new Error('XLSX 工作表结构无效');const data=child(worksheet,'sheetData');
 const cells=new Map<string,{value:string;type:SpreadsheetCellType;format?:SheetCellFormat}>(),rowIds=new Set<number>(),rowSizes=new Map<number,{size?:number;hidden?:true}>(),columnSizes=new Map<number,{size?:number;hidden?:true}>();
 // Custom row heights, hidden rows and column ranges (<cols>) become sheetDimensions; sizes are clamped to the editor's range.
 const clampSize=(axis:'row'|'column',pixels:number)=>Math.round(Math.min(SHEET_DIMENSION_LIMITS[axis].max,Math.max(SHEET_DIMENSION_LIMITS[axis].min,pixels)));
 for(const col of child(worksheet,'cols')?children(child(worksheet,'cols')!,'col'):[]){const min=Number(attribute(col,'min')),max=Number(attribute(col,'max')),width=Number(attribute(col,'width')),hidden=['1','true'].includes(attribute(col,'hidden')??''),custom=['1','true'].includes(attribute(col,'customWidth')??'');if(!Number.isInteger(min)||!Number.isInteger(max)||min<1||max<min||min>16384)continue;const value={...(custom&&Number.isFinite(width)&&width>0?{size:clampSize('column',excelWidthToPixels(width))}:{}),...(hidden?{hidden:true as const}:{})};if(!Object.keys(value).length)continue;for(let c=min-1;c<Math.min(max,min+1000,16384);c++)columnSizes.set(c,value);}let maxRow=-1,maxCol=-1,nextRow=0,textCount=0,autoCount=0,formulaCount=0,textConversions=false;
 for(const row of data?children(data,'row'):[]){const rowRef=attribute(row,'r');if(rowRef!==undefined&&!/^[1-9]\d*$/.test(rowRef))throw new Error('XLSX 行地址无效');const r=rowRef?Number(rowRef)-1:nextRow;nextRow=r+1;if(r>=10000||rowIds.has(r))throw new Error('XLSX 行数超过 1 万或行地址重复');rowIds.add(r);maxRow=Math.max(maxRow,r);let nextColumn=0;{const height=Number(attribute(row,'ht')),hidden=['1','true'].includes(attribute(row,'hidden')??''),custom=['1','true'].includes(attribute(row,'customHeight')??''),value={...(custom&&Number.isFinite(height)&&height>0?{size:clampSize('row',pointsToPixels(height))}:{}),...(hidden?{hidden:true as const}:{})};if(Object.keys(value).length)rowSizes.set(r,value);}
  for(const cell of children(row,'c')){const ref=attribute(cell,'r'),match=ref?/^([A-Z]{1,3})([1-9]\d*)$/.exec(ref):undefined;if(ref&&!match)throw new Error('XLSX 单元格地址无效');if(match&&Number(match[2])-1!==r)throw new Error('XLSX 单元格与行地址不一致');let c=nextColumn;if(match){c=0;for(const char of match[1])c=c*26+char.charCodeAt(0)-64;c--;}
   nextColumn=c+1;if(c>=16384)throw new Error('XLSX 列数超过 16384');maxCol=Math.max(maxCol,c);if((maxRow+1)*(maxCol+1)>200000)throw new Error('表格单元格超过 20 万');const key=r+':'+c;if(cells.has(key))throw new Error('XLSX 单元格地址重复');
   const formula=child(cell,'f'),cached=child(cell,'v'),kind=attribute(cell,'t');if(cached?.children.length)throw new Error('XLSX 单元格值结构无效');let value='',type:SpreadsheetCellType='auto';
   if(formula){if(formula.children.length||!formula.text||attribute(formula,'t')&&attribute(formula,'t')!=='normal'||attribute(formula,'ref')!==undefined||attribute(formula,'si')!==undefined)throw new Error('暂不支持共享、数组或数据表公式，请先在本地转换为普通公式');value='='+fromXlsxFormula(formula.text);type='formula';formulaCount++;}
   else if(kind==='s'){if(!cached||!/^\d+$/.test(cached.text)||Number(cached.text)>=strings.length)throw new Error('XLSX 共享字符串索引无效');value=strings[Number(cached.text)];type='text';}
   else if(kind==='inlineStr'){const inline=child(cell,'is');value=inline?stringValue(inline):'';type='text';}
   else if(kind==='str'||kind==='e'){value=cached?.text??'';type='text';if(kind==='e')textConversions=true;}
   else if(kind==='b'){if(!cached||!['0','1'].includes(cached.text))throw new Error('XLSX 布尔值无效');value=cached.text==='1'?'TRUE':'FALSE';type='text';textConversions=true;}
   else if(kind===undefined||kind==='n'){value=cached?.text??'';if(value&&(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(value)||!Number.isFinite(Number(value))))throw new Error('XLSX 数值无效');}
   else throw new Error('暂不支持该 XLSX 单元格类型：'+kind);
   const format=styles[Number(attribute(cell,'s')??0)];
   // Serials from the 1904 date system are rebased so every stored date uses the 1900 system.
   if(date1904&&type==='auto'&&value&&isDateNumberFormat(format?.numberFormat))value=String(Number(value)+1462);
   if(value.length>10000)throw new Error('表格单元格内容无效');cells.set(key,{value,type,...(format?{format}:{})});if(type==='text')textCount++;else if(type==='auto')autoCount++;
  }
 }
 if((maxRow+1)*(maxCol+1)>200000)throw new Error('表格单元格超过 20 万');
 const defaultType=textCount>autoCount?'text':'auto';if(Math.min(textCount,autoCount)+formulaCount>10000)throw new Error('XLSX 类型覆盖超过 1 万，请拆分工作表后导入');
 const merges=(child(worksheet,'mergeCells')?children(child(worksheet,'mergeCells')!,'mergeCell'):[]).map(merge=>{const range=cellRange(attribute(merge,'ref')??'');if(!range||range.rowSpan*range.colSpan<2||range.rowSpan>SHEET_FORMAT_LIMITS.span||range.colSpan>SHEET_FORMAT_LIMITS.span)throw new Error('XLSX 合并单元格范围无效');return range;});
 if(merges.length>SHEET_FORMAT_LIMITS.merges)throw new Error('XLSX 合并区域超过 1 万个');
 return {cells,maxRow,maxCol,defaultType,textConversions,merges,rowSizes,columnSizes};
}
export function encodeXlsxWorkbook(bytes:Uint8Array){
 const files=readOfficePackage(bytes),selected=selectedSheets(files);if(!selected.sheets.length||selected.sheets.length>MAX_WORKBOOK_SHEETS||new Set(selected.sheets.map(sheet=>sheet.path)).size!==selected.sheets.length||selected.sheets.some(sheet=>!validWorksheetName(sheet.name))||new Set(selected.sheets.map(sheet=>sheet.name.toLowerCase())).size!==selected.sheets.length)throw new Error('XLSX 工作表名称、数量或关系重复/无效');
 const styles=readStyles(files['xl/styles.xml']),workbook=files['xl/workbook.xml']?xlsxXml(files['xl/workbook.xml']):undefined,properties=workbook&&child(workbook,'workbookPr'),date1904=!!properties&&['1','true'].includes(attribute(properties,'date1904')??'');
 const shared=selected.shared?xlsxXml(selected.shared):undefined;if(shared&&shared.name!=='sst')throw new Error('XLSX 共享字符串结构无效');const strings=shared?children(shared,'si').map(stringValue):[];if(strings.some(value=>value.length>10000))throw new Error('表格单元格内容无效');
 const doc=new Y.Doc();try{
  doc.getMap('sheetCellTypes').set('$default','auto');let total=0,textConversions=false,objectCount=0,skippedRules=0,skippedCharts=0,skippedValidations=0;const dxfs=readDxfs(files['xl/styles.xml']);
  for(const [index,source]of selected.sheets.entries()){
   if(!files[source.path])throw new Error('XLSX 关系指向缺失的工作表');const parsed=parseSheet(files[source.path],strings,styles,date1904);
   // Charts placed below or beside the data need those cells to exist, so the grid grows to cover chart anchors and data.
   const objects=readSheetObjects(files,source.path,source.name,xlsxXml(files[source.path]),dxfs);
   for(const chart of objects.charts){const rows=Math.max(parsed.maxRow,chart.anchor.row,chart.range.bottom),columns=Math.max(parsed.maxCol,chart.anchor.column,chart.range.right);if(rows<10000&&columns<16384&&(rows+1)*(columns+1)<=200000){parsed.maxRow=rows;parsed.maxCol=columns;}}
   total+=(parsed.maxRow+1)*(parsed.maxCol+1);if(total>200000)throw new Error('工作簿单元格总数超过 20 万');textConversions ||= parsed.textConversions;
   const id=index===0?FIRST_WORKSHEET:'sheet-'+index,meta=new Y.Map<unknown>(),sheet=index===0?doc.getArray<Y.Array<string>>('sheet'):new Y.Array<Y.Array<string>>();meta.set('name',source.name);meta.set('defaultType',parsed.defaultType);if(index)meta.set('sheet',sheet);doc.getMap('workbookSheets').set(id,meta);doc.getArray('sheetOrder').push([id]);
   for(let r=0;r<=parsed.maxRow;r++){const row=new Y.Array<string>();row.push(Array.from({length:parsed.maxCol+1},(_,c)=>parsed.cells.get(r+':'+c)?.value??''));sheet.push([row]);}
   for(const [key,cell]of parsed.cells)if(cell.type!==parsed.defaultType){const [r,c]=key.split(':').map(Number);setSpreadsheetCellType(doc,r,c,cell.type,id);}
   const rowKeys=sheet.toArray().map(sheetRowIdentity),columnKeys=sheetColumnIds(sheet),formats=doc.getMap<SheetCellFormat>('sheetFormats'),merges=doc.getMap('sheetMerges');
   for(const [key,cell]of parsed.cells)if(cell.format){const [r,c]=key.split(':').map(Number);formats.set(sheetCellKey(id,rowKeys[r],columnKeys[c]),cell.format);}
   for(const merge of parsed.merges){if(merge.row+merge.rowSpan>rowKeys.length||merge.column+merge.colSpan>columnKeys.length)throw new Error('XLSX 合并单元格超出工作表');merges.set(sheetCellKey(id,rowKeys[merge.row],columnKeys[merge.column]),{rowSpan:merge.rowSpan,colSpan:merge.colSpan});}
   if(formats.size>SHEET_FORMAT_LIMITS.formats)throw new Error('XLSX 带格式单元格超过 5 万，请拆分工作簿');
   const dimensions=doc.getMap('sheetDimensions');for(const [r,value]of parsed.rowSizes)if(r<rowKeys.length&&dimensions.size<SHEET_DIMENSION_LIMITS.entries)dimensions.set(sheetDimensionKey(id,'row',rowKeys[r]),value);for(const [c,value]of parsed.columnSizes)if(c<columnKeys.length&&dimensions.size<SHEET_DIMENSION_LIMITS.entries)dimensions.set(sheetDimensionKey(id,'column',columnKeys[c]),value);
   // Conditional formats (clamped to the sheet) and charts; the first rule in Excel's priority order ranks highest.
   const store=doc.getMap<SheetObject>('sheetObjects'),inside=(box:{top:number;left:number;bottom:number;right:number})=>({top:box.top,left:box.left,bottom:Math.min(box.bottom,rowKeys.length-1),right:Math.min(box.right,columnKeys.length-1)});
   // Everything not stored (unsupported, over a limit, or outside the sheet) is counted for the import report.
   skippedRules+=objects.skippedRules+Math.max(0,objects.rules.length-SHEET_OBJECT_LIMITS.rulesPerSheet);skippedCharts+=objects.skippedCharts+Math.max(0,objects.charts.length-SHEET_OBJECT_LIMITS.chartsPerSheet);
   objects.rules.slice(0,SHEET_OBJECT_LIMITS.rulesPerSheet).forEach(({range,rule},i)=>{const box=inside(range);if(box.top>box.bottom||box.left>box.right||objectCount>=SHEET_OBJECT_LIMITS.objects){skippedRules++;return;}const value:SheetObject={kind:'conditional',sheet:id,range:sheetRangeKeys(sheet,box),rule,priority:objects.rules.length-i};if(validSheetObject(value)){store.set('cf-xlsx-'+index+'-'+i,value);objectCount++;}else skippedRules++;});
   objects.charts.slice(0,SHEET_OBJECT_LIMITS.chartsPerSheet).forEach((chart,i)=>{const box=inside(chart.range);if(box.top>box.bottom||box.left>box.right||chart.anchor.row>=rowKeys.length||chart.anchor.column>=columnKeys.length||objectCount>=SHEET_OBJECT_LIMITS.objects){skippedCharts++;return;}
    const value:SheetObject={kind:'chart',sheet:id,range:sheetRangeKeys(sheet,box),chartType:chart.chartType,title:chart.title,legend:chart.legend,...(chart.seriesInRows?{seriesInRows:true as const}:{}),anchor:{row:rowKeys[chart.anchor.row],column:columnKeys[chart.anchor.column]},width:chart.width,height:chart.height};
    if(validSheetObject(value)){store.set('chart-xlsx-'+index+'-'+i,value);objectCount++;}else skippedCharts++;});
   skippedValidations+=objects.skippedValidations+Math.max(0,objects.validations.length-SHEET_OBJECT_LIMITS.validationsPerSheet);
   objects.validations.slice(0,SHEET_OBJECT_LIMITS.validationsPerSheet).forEach((item,i)=>{const box=inside(item.range);if(box.top>box.bottom||box.left>box.right||objectCount>=SHEET_OBJECT_LIMITS.objects){skippedValidations++;return;}const value:SheetObject={kind:'validation',sheet:id,range:sheetRangeKeys(sheet,box),rule:item.rule,errorStyle:item.errorStyle,...(item.message?{message:item.message}:{}),...(item.prompt?{prompt:item.prompt}:{})};if(validSheetObject(value)){store.set('dv-xlsx-'+index+'-'+i,value);objectCount++;}else skippedValidations++;});
  }
  const source:XlsxSource={name:selected.sheets[0].name,sheetCount:selected.sheets.length,hasStyles:!!files['xl/styles.xml'],textConversions,...(skippedRules?{skippedRules}:{}),...(skippedCharts?{skippedCharts}:{}),...(skippedValidations?{skippedValidations}:{})};if(!validXlsxSource(source))throw new Error('XLSX 工作表信息无效');for(const [key,value]of Object.entries(source))doc.getMap('sheetXlsx').set(key,value);
  if(doc.getMap('sheetCellTypes').size>10001)throw new Error('XLSX 类型覆盖超过 1 万，请拆分工作簿');return {doc,update:Y.encodeStateAsUpdate(doc)};
 }catch(error){doc.destroy();throw error;}
}
