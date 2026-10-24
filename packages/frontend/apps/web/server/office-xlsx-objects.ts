import type * as Y from 'yjs';
import {posix} from 'node:path';
import {chartData,resolveSheetRange,sheetObjects,sheetValidations,validConditionalRule,validValidationRule,SHEET_CHART_TYPES,SHEET_OBJECT_LIMITS,type ConditionalRule,type ConditionalStyle,type SheetChart,type SheetChartType,type SheetConditionalFormat,type SheetValidationRule,type SheetValidationOperator} from '../../../../common/nbstore/src/spreadsheet-objects';
import {xlsxXml,attribute,children,type XlsxXml} from './office-xlsx-xml';
/** Office default series colours (same order as the editor's charts). */
const CHART_COLORS=['4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47','264478','9E480E','636363','997300'];

/**
 * Conditional formatting and charts for XLSX export. Charts carry cached values (c:strCache/c:numCache) so viewers
 * draw them before recalculating; anchors assume Excel's default 64 px columns and 20 px rows.
 */
const EMU=9525,COLUMN_PX=64,ROW_PX=20;
const escape=(value:string)=>value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
export function columnLetters(index:number){let name='';for(let n=index+1;n;n=Math.floor((n-1)/26))name=String.fromCharCode(65+(n-1)%26)+name;return name;}
const cell=(row:number,column:number)=>columnLetters(column)+(row+1);
const absolute=(row:number,column:number)=>'$'+columnLetters(column)+'$'+(row+1);
const sheetRef=(name:string)=>"'"+name.replaceAll("'","''")+"'";
const number=(text:string)=>{const value=Number(text.trim());return text.trim()!==''&&Number.isFinite(value)?value:undefined;};
/** Formula operand for a rule value: numbers as-is, anything else as a quoted string. */
const operand=(value:string)=>number(value)!==undefined?String(number(value)):'"'+value.replaceAll('"','""')+'"';

/** Differential formats (dxfs) shared by all sheets; index is the dxfId. */
export function dxfRegistry(){
 const list:string[]=[];
 return {id(style:ConditionalStyle){const xml='<dxf>'+(style.bold||style.color?`<font>${style.bold?'<b/>':''}${style.color?`<color rgb="FF${style.color}"/>`:''}</font>`:'')+(style.fill?`<fill><patternFill patternType="solid"><fgColor rgb="FF${style.fill}"/><bgColor rgb="FF${style.fill}"/></patternFill></fill>`:'')+'</dxf>';const at=list.indexOf(xml);if(at>=0)return at;list.push(xml);return list.length-1;},
  xml(){return list.length?`<dxfs count="${list.length}">${list.join('')}</dxfs>`:'';}};
}
function ruleXml(rule:ConditionalRule,priority:number,top:string,dxf:(style:ConditionalStyle)=>number){
 const p=`priority="${priority}"`,d=(style:ConditionalStyle)=>`dxfId="${dxf(style)}"`;
 switch(rule.type){
  case 'greaterThan':case 'lessThan':case 'equal':case 'notEqual':case 'greaterOrEqual':case 'lessOrEqual':{
   const op={greaterThan:'greaterThan',lessThan:'lessThan',equal:'equal',notEqual:'notEqual',greaterOrEqual:'greaterThanOrEqual',lessOrEqual:'lessThanOrEqual'}[rule.type];
   return `<cfRule type="cellIs" ${d(rule.style)} ${p} operator="${op}"><formula>${escape(operand(rule.value))}</formula></cfRule>`;}
  case 'between':case 'notBetween':return `<cfRule type="cellIs" ${d(rule.style)} ${p} operator="${rule.type}"><formula>${escape(operand(rule.value))}</formula><formula>${escape(operand(rule.value2))}</formula></cfRule>`;
  case 'textContains':return `<cfRule type="containsText" ${d(rule.style)} ${p} operator="containsText" text="${escape(rule.value)}"><formula>${escape(`NOT(ISERROR(SEARCH(${operand(rule.value)},${top})))`)}</formula></cfRule>`;
  case 'textNotContains':return `<cfRule type="notContainsText" ${d(rule.style)} ${p} operator="notContains" text="${escape(rule.value)}"><formula>${escape(`ISERROR(SEARCH(${operand(rule.value)},${top}))`)}</formula></cfRule>`;
  case 'beginsWith':return `<cfRule type="beginsWith" ${d(rule.style)} ${p} operator="beginsWith" text="${escape(rule.value)}"><formula>${escape(`LEFT(${top},LEN(${operand(rule.value)}))=${operand(rule.value)}`)}</formula></cfRule>`;
  case 'endsWith':return `<cfRule type="endsWith" ${d(rule.style)} ${p} operator="endsWith" text="${escape(rule.value)}"><formula>${escape(`RIGHT(${top},LEN(${operand(rule.value)}))=${operand(rule.value)}`)}</formula></cfRule>`;
  case 'duplicate':return `<cfRule type="duplicateValues" ${d(rule.style)} ${p}/>`;
  case 'unique':return `<cfRule type="uniqueValues" ${d(rule.style)} ${p}/>`;
  case 'blank':return `<cfRule type="containsBlanks" ${d(rule.style)} ${p}><formula>${escape(`LEN(TRIM(${top}))=0`)}</formula></cfRule>`;
  case 'notBlank':return `<cfRule type="notContainsBlanks" ${d(rule.style)} ${p}><formula>${escape(`LEN(TRIM(${top}))>0`)}</formula></cfRule>`;
  case 'top':case 'bottom':return `<cfRule type="top10" ${d(rule.style)} ${p} rank="${rule.count}"${rule.percent?' percent="1"':''}${rule.type==='bottom'?' bottom="1"':''}/>`;
  case 'aboveAverage':return `<cfRule type="aboveAverage" ${d(rule.style)} ${p}/>`;
  case 'belowAverage':return `<cfRule type="aboveAverage" ${d(rule.style)} ${p} aboveAverage="0"/>`;
  case 'colorScale':return `<cfRule type="colorScale" ${p}><colorScale><cfvo type="min"/>${rule.colors.length===3?'<cfvo type="percentile" val="50"/>':''}<cfvo type="max"/>${rule.colors.map(color=>`<color rgb="FF${color}"/>`).join('')}</colorScale></cfRule>`;
  case 'dataBar':return `<cfRule type="dataBar" ${p}><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF${rule.color}"/></dataBar></cfRule>`;
 }
}
/** `<conditionalFormatting>` elements for one sheet, highest priority first (priority 1 in Excel). */
export function conditionalXml(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>,dxf:(style:ConditionalStyle)=>number){
 const rules=sheetObjects(doc,sheetId).flatMap(([,object])=>object.kind==='conditional'?[object as SheetConditionalFormat]:[]).sort((a,b)=>b.priority-a.priority);let priority=0;
 return rules.map(rule=>{const range=resolveSheetRange(sheet,rule.range);if(!range)return '';const ref=cell(range.top,range.left)+(range.bottom>range.top||range.right>range.left?':'+cell(range.bottom,range.right):'');
  return `<conditionalFormatting sqref="${ref}">${ruleXml(rule.rule,++priority,cell(range.top,range.left),dxf)}</conditionalFormatting>`;}).join('');
}

const VALIDATION_OPERATORS:Record<SheetValidationOperator,string>={between:'between',notBetween:'notBetween',equal:'equal',notEqual:'notEqual',greaterThan:'greaterThan',lessThan:'lessThan',greaterOrEqual:'greaterThanOrEqual',lessOrEqual:'lessThanOrEqual'};
/** `<dataValidations>` for one sheet: literal lists as a quoted formula, comparisons with operator and bounds. */
export function validationXml(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>){
 const items=sheetValidations(doc,sheetId).flatMap(([,validation])=>{
  const range=resolveSheetRange(sheet,validation.range);if(!range)return [];const ref=cell(range.top,range.left)+(range.bottom>range.top||range.right>range.left?':'+cell(range.bottom,range.right):''),rule=validation.rule;
  const messages=`${validation.message?` errorTitle="输入值无效" error="${escape(validation.message)}"`:''}${validation.prompt?` promptTitle="提示" prompt="${escape(validation.prompt)}"`:''}`;
  const formulas=rule.type==='list'?`<formula1>${escape('"'+rule.items.join(',').replaceAll('"','""')+'"')}</formula1>`:`<formula1>${escape(rule.value)}</formula1>${rule.value2!==undefined?`<formula2>${escape(rule.value2)}</formula2>`:''}`;
  return [`<dataValidation type="${rule.type}"${rule.type!=='list'?` operator="${VALIDATION_OPERATORS[rule.operator]}"`:''} errorStyle="${validation.errorStyle}" allowBlank="1" showInputMessage="1" showErrorMessage="1"${messages} sqref="${ref}">${formulas}</dataValidation>`];
 });
 return items.length?`<dataValidations count="${items.length}">${items.join('')}</dataValidations>`:'';
}

/** One chart part plus its anchor in the sheet's drawing. */
function chartXml(chart:SheetChart,sheetName:string,sheet:Y.Array<Y.Array<string>>,value:(r:number,c:number)=>string,label:(r:number,c:number)=>string){
 const range=resolveSheetRange(sheet,chart.range);if(!range)return;const data=chartData(range,value,chart.seriesInRows,label);if(!data.series.length||!data.categories.length)return;
 const r0=range.top+(data.headerRow?1:0),c0=range.left+(data.headerColumn?1:0),name=sheetRef(sheetName);
 const strCache=(values:string[])=>`<c:strCache><c:ptCount val="${values.length}"/>${values.map((v,i)=>`<c:pt idx="${i}"><c:v>${escape(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
 const numCache=(values:(number|null)[])=>`<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v,i)=>v===null?'':`<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
 const points=data.categories.length,categoryRef=chart.seriesInRows?(data.headerRow?`${name}!${absolute(range.top,c0)}:${absolute(range.top,c0+points-1)}`:''):(data.headerColumn?`${name}!${absolute(r0,range.left)}:${absolute(r0+points-1,range.left)}`:'');
 const series=(chart.chartType==='pie'?data.series.slice(0,1):data.series).map((item,s)=>{
  const nameRef=chart.seriesInRows?(data.headerColumn?`${name}!${absolute(r0+s,range.left)}`:''):(data.headerRow?`${name}!${absolute(range.top,c0+s)}`:'');
  const valueRef=chart.seriesInRows?`${name}!${absolute(r0+s,c0)}:${absolute(r0+s,c0+points-1)}`:`${name}!${absolute(r0,c0+s)}:${absolute(r0+points-1,c0+s)}`;
  // Explicit colours: without a theme part, readers otherwise fall back to no fill (LibreOffice draws empty charts).
  const color=CHART_COLORS[s%CHART_COLORS.length],fill=(hex:string)=>`<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  const shape=chart.chartType==='line'?`<c:spPr><a:ln w="28575" cap="rnd">${fill(color)}<a:round/></a:ln></c:spPr>`:`<c:spPr>${fill(color)}${chart.chartType==='pie'?'<a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>':''}</c:spPr>`;
  const slices=chart.chartType==='pie'?item.values.map((_,i)=>`<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fill(CHART_COLORS[i%CHART_COLORS.length])}<a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join(''):'';
  return `<c:ser><c:idx val="${s}"/><c:order val="${s}"/>${nameRef?`<c:tx><c:strRef><c:f>${escape(nameRef)}</c:f>${strCache([item.name])}</c:strRef></c:tx>`:`<c:tx><c:v>${escape(item.name)}</c:v></c:tx>`}`+shape
   +(chart.chartType==='line'?`<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(color)}</c:spPr></c:marker>`:'')+(chart.chartType==='bar'||chart.chartType==='column'?'<c:invertIfNegative val="0"/>':'')+slices
   +(categoryRef?`<c:cat><c:strRef><c:f>${escape(categoryRef)}</c:f>${strCache(data.categories)}</c:strRef></c:cat>`:'')
   +`<c:val><c:numRef><c:f>${escape(valueRef)}</c:f>${numCache(item.values)}</c:numRef></c:val>`+(chart.chartType==='line'?'<c:smooth val="0"/>':'')+'</c:ser>';
 }).join('');
 const axes='<c:axId val="500000001"/><c:axId val="500000002"/>',bar=chart.chartType==='bar';
 const group=chart.chartType==='pie'?`<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`
  :chart.chartType==='line'?`<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axes}</c:lineChart>`
  :chart.chartType==='area'?`<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${axes}</c:areaChart>`
  :`<c:barChart><c:barDir val="${bar?'bar':'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/>${axes}</c:barChart>`;
 const axisXml=chart.chartType==='pie'?'':`<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${bar?'l':'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`
  +`<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${bar?'b':'l'}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
 const title=chart.title?`<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${escape(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`:'<c:autoTitleDeleted val="1"/>';
 const xml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>${title}<c:plotArea><c:layout/>${group}${axisXml}</c:plotArea>${chart.legend?'<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>':''}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
 return xml;
}
/**
 * Drawing + chart parts for one sheet. `chartNumber` numbers chart parts across the workbook. Returns files, the
 * sheet relationship to add and the content-type overrides.
 */
export function sheetCharts(doc:Y.Doc,sheetId:string,sheetName:string,sheet:Y.Array<Y.Array<string>>,drawingNumber:number,chartNumber:number,value:(r:number,c:number)=>string,label:(r:number,c:number)=>string){
 const files:Record<string,string>={},types:string[]=[],anchors:string[]=[],relations:string[]=[];let next=chartNumber;
 for(const [,object]of sheetObjects(doc,sheetId)){
  if(object.kind!=='chart')continue;const chart=object as SheetChart,xml=chartXml(chart,sheetName,sheet,value,label);if(!xml)continue;
  const anchor=resolveSheetRange(sheet,{top:chart.anchor.row,bottom:chart.anchor.row,left:chart.anchor.column,right:chart.anchor.column});if(!anchor)continue;
  const n=next++,rel='rIdChart'+n,toColumn=anchor.left+Math.floor(chart.width/COLUMN_PX),toRow=anchor.top+Math.floor(chart.height/ROW_PX);
  files[`xl/charts/chart${n}.xml`]=xml;types.push(`<Override PartName="/xl/charts/chart${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
  relations.push(`<Relationship Id="${rel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${n}.xml"/>`);
  anchors.push(`<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${anchor.left}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.top}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toColumn}</xdr:col><xdr:colOff>${(chart.width%COLUMN_PX)*EMU}</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>${(chart.height%ROW_PX)*EMU}</xdr:rowOff></xdr:to>`
   +`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${n+1}" name="${escape(chart.title||'图表 '+n)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rel}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`);
 }
 if(!anchors.length)return;
 files[`xl/drawings/drawing${drawingNumber}.xml`]=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors.join('')}</xdr:wsDr>`;
 files[`xl/drawings/_rels/drawing${drawingNumber}.xml.rels`]=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relations.join('')}</Relationships>`;
 types.push(`<Override PartName="/xl/drawings/drawing${drawingNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
 return {files,types,relation:`<Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNumber}.xml"/>`,next};
}

// ---- import ---------------------------------------------------------------------------------------------------------
/** Lenient child lookup (charts often repeat elements that the strict `child` helper rejects). */
const first=(node:XlsxXml|undefined,name:string)=>node?.children.find(item=>item.name===name);
const deep=(node:XlsxXml|undefined,path:string[])=>path.reduce<XlsxXml|undefined>((current,name)=>first(current,name),node);
const rgbOf=(node:XlsxXml|undefined)=>{const value=node&&attribute(node,'rgb');return value&&/^[0-9A-Fa-f]{8}$/.test(value)?value.slice(2).toUpperCase():value&&/^[0-9A-Fa-f]{6}$/.test(value)?value.toUpperCase():undefined;};
type Box={top:number;left:number;bottom:number;right:number};
const lettersToColumn=(letters:string)=>[...letters.toUpperCase()].reduce((n,char)=>n*26+char.charCodeAt(0)-64,0)-1;
/** "A1" or "A1:C9" (with optional $) as a box; whole rows/columns and other forms are not supported. */
export function parseBox(ref:string):Box|undefined{
 const match=/^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(ref.trim());if(!match)return;
 const r1=Number(match[2])-1,c1=lettersToColumn(match[1]),r2=match[4]?Number(match[4])-1:r1,c2=match[3]?lettersToColumn(match[3]):c1;
 return {top:Math.min(r1,r2),bottom:Math.max(r1,r2),left:Math.min(c1,c2),right:Math.max(c1,c2)};
}
/** Differential formats from styles.xml, as highlight styles (theme/indexed colours are not carried over). */
export function readDxfs(bytes:Uint8Array|undefined):(ConditionalStyle|undefined)[]{
 if(!bytes)return [];const root=xlsxXml(bytes),list=first(root,'dxfs');
 return (list?children(list,'dxf'):[]).map(dxf=>{const font=first(dxf,'font'),pattern=deep(dxf,['fill','patternFill']),bold=first(font,'b'),fill=rgbOf(first(pattern,'bgColor'))??rgbOf(first(pattern,'fgColor')),color=rgbOf(first(font,'color'));
  const style:ConditionalStyle={...(fill?{fill}:{}),...(color?{color}:{}),...(bold&&!['0','false'].includes(attribute(bold,'val')??'1')?{bold:true}:{})};return Object.keys(style).length?style:undefined;});
}
const unquote=(formula:string|undefined)=>{if(formula===undefined)return;const text=formula.trim();if(/^"(?:[^"]|"")*"$/.test(text))return text.slice(1,-1).replaceAll('""','"');if(text!==''&&Number.isFinite(Number(text)))return text;};
function ruleFrom(node:XlsxXml,dxfs:(ConditionalStyle|undefined)[]):ConditionalRule|undefined{
 const type=attribute(node,'type'),dxf=attribute(node,'dxfId'),style=dxf!==undefined?dxfs[Number(dxf)]:undefined,formulas=children(node,'formula').map(item=>item.text),operator=attribute(node,'operator');
 let rule:Record<string,unknown>|undefined;
 if(type==='cellIs'&&style){
  const value=unquote(formulas[0]),value2=unquote(formulas[1]);
  const ops:Record<string,string>={greaterThan:'greaterThan',lessThan:'lessThan',equal:'equal',notEqual:'notEqual',greaterThanOrEqual:'greaterOrEqual',lessThanOrEqual:'lessOrEqual',between:'between',notBetween:'notBetween'};
  if(operator&&ops[operator]&&value!==undefined)rule=operator==='between'||operator==='notBetween'?(value2===undefined?undefined:{type:ops[operator],value,value2,style}):{type:ops[operator],value,style};
 }
 else if(['containsText','notContainsText','beginsWith','endsWith'].includes(type??'')&&style){const text=attribute(node,'text');if(text)rule={type:{containsText:'textContains',notContainsText:'textNotContains',beginsWith:'beginsWith',endsWith:'endsWith'}[type!],value:text,style};}
 else if(type==='duplicateValues'&&style)rule={type:'duplicate',style};else if(type==='uniqueValues'&&style)rule={type:'unique',style};
 else if(type==='containsBlanks'&&style)rule={type:'blank',style};else if(type==='notContainsBlanks'&&style)rule={type:'notBlank',style};
 else if(type==='top10'&&style){const rank=Number(attribute(node,'rank')??10);rule={type:['1','true'].includes(attribute(node,'bottom')??'')?'bottom':'top',count:rank,...(['1','true'].includes(attribute(node,'percent')??'')?{percent:true}:{}),style};}
 else if(type==='aboveAverage'&&style&&!attribute(node,'stdDev')&&!['1','true'].includes(attribute(node,'equalAverage')??''))rule={type:['0','false'].includes(attribute(node,'aboveAverage')??'1')?'belowAverage':'aboveAverage',style};
 else if(type==='colorScale'){const colors=children(first(node,'colorScale')??node,'color').map(rgbOf);if(colors.every(Boolean)&&[2,3].includes(colors.length))rule={type:'colorScale',colors};}
 else if(type==='dataBar'){const color=rgbOf(first(first(node,'dataBar'),'color'));if(color)rule={type:'dataBar',color};}
 return rule&&validConditionalRule(rule)?rule as ConditionalRule:undefined;
}
export type ImportedChart={chartType:SheetChartType;title:string;legend:boolean;seriesInRows:boolean;range:Box;anchor:{row:number;column:number};width:number;height:number};
export type ImportedValidation={range:Box;rule:SheetValidationRule;errorStyle:'stop'|'warning';message?:string;prompt?:string};
export type ImportedSheetObjects={rules:{range:Box;rule:ConditionalRule}[];charts:ImportedChart[];validations:ImportedValidation[];skippedRules:number;skippedCharts:number;skippedValidations:number};
const resolve=(from:string,target:string)=>{if(/[\\:#?]/.test(target))return;const path=target.startsWith('/')?target.slice(1):posix.normalize(posix.join(posix.dirname(from),target));return path.startsWith('xl/')&&!path.split('/').includes('..')?path:undefined;};
const relsFor=(path:string)=>posix.join(posix.dirname(path),'_rels',posix.basename(path)+'.rels');
function relations(files:Record<string,Uint8Array>,path:string){const bytes=files[relsFor(path)],map=new Map<string,{type:string;target:string}>();if(!bytes)return map;const root=xlsxXml(bytes);
 for(const item of children(root,'Relationship')){const id=attribute(item,'Id'),target=attribute(item,'Target'),type=attribute(item,'Type');if(id&&target&&type&&attribute(item,'TargetMode')!=='External'){const resolved=resolve(path,target);if(resolved)map.set(id,{type,target:resolved});}}return map;}
/** A series formula like 'Sheet 1'!$B$2:$B$9 on `sheetName`; other sheets and named ranges are not supported. */
function seriesBox(formula:string|undefined,sheetName:string){
 if(!formula)return;const match=/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(formula.trim());if(!match)return;const name=(match[1]??match[2]).replaceAll("''","'");
 return name.toLowerCase()===sheetName.toLowerCase()?parseBox(match[3]):undefined;
}
const union=(boxes:Box[])=>({top:Math.min(...boxes.map(b=>b.top)),left:Math.min(...boxes.map(b=>b.left)),bottom:Math.max(...boxes.map(b=>b.bottom)),right:Math.max(...boxes.map(b=>b.right))});
function chartFrom(bytes:Uint8Array,sheetName:string):Omit<ImportedChart,'anchor'|'width'|'height'>|undefined{
 const space=xlsxXml(bytes),chart=first(space,'chart'),plot=first(chart,'plotArea');if(!plot)return;
 const kinds:Record<string,SheetChartType>={barChart:'column',bar3DChart:'column',lineChart:'line',line3DChart:'line',pieChart:'pie',pie3DChart:'pie',doughnutChart:'pie',areaChart:'area',area3DChart:'area'};
 const groups=plot.children.filter(item=>kinds[item.name]);if(groups.length!==1)return;const group=groups[0];let type=kinds[group.name];
 if(type==='column'&&attribute(first(group,'barDir')??group,'val')==='bar')type='bar';if(!SHEET_CHART_TYPES.includes(type))return;
 const series=children(group,'ser');if(!series.length||series.length>SHEET_OBJECT_LIMITS.series)return;
 const values=series.map(ser=>seriesBox(deep(ser,['val','numRef','f'])?.text,sheetName));if(values.some(box=>!box))return;const boxes=values as Box[];
 const inColumns=boxes.every(box=>box.left===box.right),inRows=boxes.every(box=>box.top===box.bottom);if(!inColumns&&!inRows)return;
 const extra=series.flatMap(ser=>[seriesBox((deep(ser,['cat','strRef','f'])??deep(ser,['cat','numRef','f']))?.text,sheetName),seriesBox(deep(ser,['tx','strRef','f'])?.text,sheetName)]).filter((box):box is Box=>!!box);
 const titleRuns=deep(chart,['title','tx','rich']);const title=titleRuns?titleRuns.children.filter(item=>item.name==='p').map(p=>children(p,'r').map(run=>first(run,'t')?.text??'').join('')).join(' ').trim():'';
 return {chartType:type,title:title.slice(0,SHEET_OBJECT_LIMITS.title),legend:!!first(chart,'legend'),seriesInRows:!inColumns,range:union([...boxes,...extra])};
}
/** Conditional formats and charts of one worksheet; unsupported ones are counted, never fail the import. */
export function readSheetObjects(files:Record<string,Uint8Array>,sheetPath:string,sheetName:string,worksheet:XlsxXml,dxfs:(ConditionalStyle|undefined)[]):ImportedSheetObjects{
 const result:ImportedSheetObjects={rules:[],charts:[],validations:[],skippedRules:0,skippedCharts:0,skippedValidations:0};
 // Literal lists and number/text-length rules (Excel only blocks input when showErrorMessage is on and the style is stop); range-based lists, dates, times and custom formulas are counted as skipped.
 const operators=Object.fromEntries(Object.entries(VALIDATION_OPERATORS).map(([key,value])=>[value,key])) as Record<string,SheetValidationOperator>;
 for(const node of children(first(worksheet,'dataValidations')??worksheet,'dataValidation')){
  const type=attribute(node,'type'),boxes=(attribute(node,'sqref')??'').split(/\s+/).filter(Boolean).map(parseBox),f1=first(node,'formula1')?.text,f2=first(node,'formula2')?.text;let rule:Record<string,unknown>|undefined;
  if(type==='list'&&f1&&/^"(?:[^"]|"")*"$/.test(f1.trim()))rule={type:'list',items:[...new Set(f1.trim().slice(1,-1).replaceAll('""','"').split(',').map(item=>item.trim()).filter(Boolean))]};
  else if(['whole','decimal','textLength'].includes(type??'')){const operator=operators[attribute(node,'operator')??'between'];rule=operator==='between'||operator==='notBetween'?{type,operator,value:f1?.trim(),value2:f2?.trim()}:{type,operator,value:f1?.trim()};}
  const text=(name:string)=>{const value=attribute(node,name);return value&&value.length<=SHEET_OBJECT_LIMITS.text&&!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value)?value:undefined;},message=text('error'),prompt=text('prompt');
  if(!rule||!validValidationRule(rule)||!boxes.length||boxes.some(box=>!box)){result.skippedValidations++;continue;}
  for(const box of boxes)result.validations.push({range:box!,rule:rule as SheetValidationRule,errorStyle:['1','true'].includes(attribute(node,'showErrorMessage')??'0')&&(attribute(node,'errorStyle')??'stop')==='stop'?'stop':'warning',...(message?{message}:{}),...(prompt?{prompt}:{})});
 }
 const formats=children(worksheet,'conditionalFormatting').flatMap(format=>children(format,'cfRule').map(rule=>({ranges:(attribute(format,'sqref')??'').split(/\s+/).filter(Boolean),rule,priority:Number(attribute(rule,'priority')??1e6)}))).sort((a,b)=>a.priority-b.priority);
 for(const {ranges,rule}of formats){const parsed=ruleFrom(rule,dxfs),boxes=ranges.map(parseBox);if(!parsed||!boxes.length||boxes.some(box=>!box)){result.skippedRules++;continue;}for(const box of boxes)result.rules.push({range:box!,rule:parsed});}
 try{
  const drawing=[...relations(files,sheetPath)].find(([id,item])=>item.type.endsWith('/drawing')&&children(worksheet,'drawing').some(node=>attribute(node,'id')===id))?.[1];
  if(drawing&&files[drawing.target]){const root=xlsxXml(files[drawing.target]),links=relations(files,drawing.target);
   for(const anchor of root.children.filter(item=>item.name==='twoCellAnchor'||item.name==='oneCellAnchor')){
    const ref=deep(anchor,['graphicFrame','graphic','graphicData','chart']),id=ref&&attribute(ref,'id'),part=id?links.get(id):undefined;if(!part){if(ref)result.skippedCharts++;continue;}
    const from=first(anchor,'from'),to=first(anchor,'to'),num=(node:XlsxXml|undefined,name:string)=>Number(first(node,name)?.text??0);
    const parsed=files[part.target]?chartFrom(files[part.target],sheetName):undefined;if(!parsed||!from){result.skippedCharts++;continue;}
    const ext=first(anchor,'ext'),width=to?(num(to,'col')-num(from,'col'))*COLUMN_PX+(num(to,'colOff')-num(from,'colOff'))/EMU:Number(ext&&attribute(ext,'cx'))/EMU,height=to?(num(to,'row')-num(from,'row'))*ROW_PX+(num(to,'rowOff')-num(from,'rowOff'))/EMU:Number(ext&&attribute(ext,'cy'))/EMU;
    const clamp=(value:number,fallback:number)=>Math.round(Math.min(2000,Math.max(160,Number.isFinite(value)&&value>0?value:fallback)));
    result.charts.push({...parsed,anchor:{row:num(from,'row'),column:num(from,'col')},width:clamp(width,480),height:clamp(height,288)});
   }}
 }catch{result.skippedCharts++;}
 return result;
}
