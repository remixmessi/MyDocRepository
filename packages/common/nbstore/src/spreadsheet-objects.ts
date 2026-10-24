import * as Y from 'yjs';
import {sheetColumnIds,sheetRowIdentity} from './sheet-columns';

/**
 * Worksheet objects: conditional formatting rules and charts, kept in the `sheetObjects` root keyed by object id.
 * Ranges use stable row/column identities so they follow row and column inserts and deletes; values are plain JSON
 * replaced as a whole, with a revision of the stored JSON to detect concurrent edits.
 */
export type SheetRangeKeys={top:string;left:string;bottom:string;right:string};
export type ConditionalStyle={fill?:string;color?:string;bold?:true};
export type ConditionalRule=
 {type:'greaterThan'|'lessThan'|'equal'|'notEqual'|'greaterOrEqual'|'lessOrEqual';value:string;style:ConditionalStyle}|
 {type:'between'|'notBetween';value:string;value2:string;style:ConditionalStyle}|
 {type:'textContains'|'textNotContains'|'beginsWith'|'endsWith';value:string;style:ConditionalStyle}|
 {type:'duplicate'|'unique'|'blank'|'notBlank';style:ConditionalStyle}|
 {type:'top'|'bottom';count:number;percent?:true;style:ConditionalStyle}|
 {type:'aboveAverage'|'belowAverage';style:ConditionalStyle}|
 {type:'colorScale';colors:[string,string]|[string,string,string]}|
 {type:'dataBar';color:string};
export type SheetConditionalFormat={kind:'conditional';sheet:string;range:SheetRangeKeys;rule:ConditionalRule;priority:number};
export type SheetChartType='column'|'bar'|'line'|'pie'|'area';
export type SheetChart={kind:'chart';sheet:string;range:SheetRangeKeys;chartType:SheetChartType;title:string;seriesInRows?:true;legend:boolean;anchor:{row:string;column:string};width:number;height:number};
/** Data validation (数据有效性): a literal dropdown list or a number/text-length comparison. */
export type SheetValidationOperator='between'|'notBetween'|'equal'|'notEqual'|'greaterThan'|'lessThan'|'greaterOrEqual'|'lessOrEqual';
export type SheetValidationRule={type:'list';items:string[]}|{type:'whole'|'decimal'|'textLength';operator:SheetValidationOperator;value:string;value2?:string};
export type SheetValidation={kind:'validation';sheet:string;range:SheetRangeKeys;rule:SheetValidationRule;errorStyle:'stop'|'warning';message?:string;prompt?:string};
export type SheetObject=SheetConditionalFormat|SheetChart|SheetValidation;
export const SHEET_OBJECT_LIMITS={objects:1000,rulesPerSheet:200,chartsPerSheet:50,validationsPerSheet:200,text:255,title:200,series:20,points:500,listItems:100,listText:255};
export const SHEET_VALIDATION_OPERATORS:SheetValidationOperator[]=['between','notBetween','equal','notEqual','greaterThan','lessThan','greaterOrEqual','lessOrEqual'];
export const SHEET_CHART_TYPES:SheetChartType[]=['column','bar','line','pie','area'];
const COMPARE=['greaterThan','lessThan','equal','notEqual','greaterOrEqual','lessOrEqual'],BETWEEN=['between','notBetween'],TEXT=['textContains','textNotContains','beginsWith','endsWith'],PLAIN=['duplicate','unique','blank','notBlank','aboveAverage','belowAverage'];

const record=(value:unknown):value is Record<string,unknown>=>!!value&&typeof value==='object'&&!Array.isArray(value)&&!(value instanceof Y.AbstractType);
const exact=(value:Record<string,unknown>,required:string[],optional:string[]=[])=>required.every(key=>key in value)&&Object.keys(value).every(key=>required.includes(key)||optional.includes(key));
const hex=(value:unknown)=>typeof value==='string'&&/^[0-9A-F]{6}$/.test(value);
const rowKey=(value:unknown)=>typeof value==='string'&&/^\d+-\d+$/.test(value);
const columnKey=(value:unknown)=>typeof value==='string'&&/^[a-zA-Z0-9-]{1,180}$/.test(value);
const shortText=(value:unknown)=>typeof value==='string'&&value.length<=SHEET_OBJECT_LIMITS.text&&!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value);
export const validSheetObjectId=(value:unknown):value is string=>typeof value==='string'&&/^[a-zA-Z0-9-]{8,80}$/.test(value);
function validRange(value:unknown):value is SheetRangeKeys{return record(value)&&exact(value,['top','left','bottom','right'])&&rowKey(value.top)&&rowKey(value.bottom)&&columnKey(value.left)&&columnKey(value.right);}
function validStyle(value:unknown):value is ConditionalStyle{return record(value)&&Object.keys(value).length>0&&exact(value,[],['fill','color','bold'])&&(value.fill===undefined||hex(value.fill))&&(value.color===undefined||hex(value.color))&&(value.bold===undefined||value.bold===true);}
export function validConditionalRule(value:unknown):value is ConditionalRule{
 if(!record(value)||typeof value.type!=='string')return false;const type=value.type;
 if(COMPARE.includes(type)||TEXT.includes(type))return exact(value,['type','value','style'])&&shortText(value.value)&&(value.value as string).length>0&&validStyle(value.style);
 if(BETWEEN.includes(type))return exact(value,['type','value','value2','style'])&&shortText(value.value)&&shortText(value.value2)&&(value.value as string).length>0&&(value.value2 as string).length>0&&validStyle(value.style);
 if(PLAIN.includes(type))return exact(value,['type','style'])&&validStyle(value.style);
 if(type==='top'||type==='bottom')return exact(value,['type','count','style'],['percent'])&&Number.isInteger(value.count)&&Number(value.count)>=1&&Number(value.count)<=(value.percent?100:1000)&&(value.percent===undefined||value.percent===true)&&validStyle(value.style);
 if(type==='colorScale')return exact(value,['type','colors'])&&Array.isArray(value.colors)&&[2,3].includes(value.colors.length)&&value.colors.every(hex);
 return type==='dataBar'&&exact(value,['type','color'])&&hex(value.color);
}
/** Excel stores a literal list as one quoted, comma-separated formula of at most 255 characters. */
export function validValidationRule(value:unknown):value is SheetValidationRule{
 if(!record(value))return false;
 if(value.type==='list')return exact(value,['type','items'])&&Array.isArray(value.items)&&value.items.length>=1&&value.items.length<=SHEET_OBJECT_LIMITS.listItems&&value.items.every(item=>typeof item==='string'&&item.trim()!==''&&item===item.trim()&&!item.includes(',')&&shortText(item))&&new Set(value.items).size===value.items.length&&value.items.join(',').length<=SHEET_OBJECT_LIMITS.listText;
 if(!['whole','decimal','textLength'].includes(value.type as string)||!SHEET_VALIDATION_OPERATORS.includes(value.operator as SheetValidationOperator))return false;
 const between=value.operator==='between'||value.operator==='notBetween',ok=(text:unknown)=>typeof text==='string'&&text.trim()!==''&&Number.isFinite(Number(text))&&(value.type!=='decimal'?Number.isInteger(Number(text)):true)&&(value.type!=='textLength'||Number(text)>=0);
 return exact(value,between?['type','operator','value','value2']:['type','operator','value'])&&ok(value.value)&&(!between||ok(value.value2));
}
export function validSheetObject(value:unknown):value is SheetObject{
 if(!record(value)||typeof value.sheet!=='string'||!/^[a-zA-Z0-9-]{1,128}$/.test(value.sheet)||!validRange(value.range))return false;
 if(value.kind==='conditional')return exact(value,['kind','sheet','range','rule','priority'])&&validConditionalRule(value.rule)&&typeof value.priority==='number'&&Number.isFinite(value.priority)&&Math.abs(value.priority)<=1e9;
 if(value.kind==='validation')return exact(value,['kind','sheet','range','rule','errorStyle'],['message','prompt'])&&validValidationRule(value.rule)&&(value.errorStyle==='stop'||value.errorStyle==='warning')&&[value.message,value.prompt].every(text=>text===undefined||shortText(text)&&(text as string).length>0);
 return value.kind==='chart'&&exact(value,['kind','sheet','range','chartType','title','legend','anchor','width','height'],['seriesInRows'])&&SHEET_CHART_TYPES.includes(value.chartType as SheetChartType)
  &&typeof value.title==='string'&&value.title.length<=SHEET_OBJECT_LIMITS.title&&!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value.title)&&typeof value.legend==='boolean'&&(value.seriesInRows===undefined||value.seriesInRows===true)
  &&record(value.anchor)&&exact(value.anchor,['row','column'])&&rowKey(value.anchor.row)&&columnKey(value.anchor.column)
  &&[value.width,value.height].every(size=>Number.isInteger(size)&&Number(size)>=160&&Number(size)<=2000);
}
/** Revision of an object for optimistic concurrency: its stored JSON (a missing object is ""). */
export const sheetObjectRevision=(value:SheetObject|undefined)=>value?JSON.stringify(value):'';
export function sheetObjects(doc:Y.Doc,sheetId?:string){
 if(!doc.share.has('sheetObjects'))return [] as [string,SheetObject][];
 return [...doc.getMap<SheetObject>('sheetObjects')].filter((entry):entry is [string,SheetObject]=>validSheetObject(entry[1])&&(sheetId===undefined||entry[1].sheet===sheetId));
}
export type ResolvedRange={top:number;left:number;bottom:number;right:number};
/** Current indexes of a stored range; undefined when a corner row/column no longer exists. */
export function resolveSheetRange(sheet:Y.Array<Y.Array<string>>,range:SheetRangeKeys):ResolvedRange|undefined{
 const rows=new Map(sheet.toArray().map((row,i)=>[sheetRowIdentity(row),i])),columns=sheetColumnIds(sheet),a=rows.get(range.top),b=rows.get(range.bottom),c=columns.indexOf(range.left),d=columns.indexOf(range.right);
 if(a===undefined||b===undefined||c<0||d<0)return;return {top:Math.min(a,b),bottom:Math.max(a,b),left:Math.min(c,d),right:Math.max(c,d)};
}
export function sheetRangeKeys(sheet:Y.Array<Y.Array<string>>,range:ResolvedRange):SheetRangeKeys{
 const rows=sheet.toArray(),columns=sheetColumnIds(sheet);return {top:sheetRowIdentity(rows[range.top]),bottom:sheetRowIdentity(rows[range.bottom]),left:columns[range.left],right:columns[range.right]};
}
/**
 * Keep object ranges valid across a row/column delete (call before mutating the grid): a deleted corner moves inward
 * to the next row/column of the range, and an object whose whole range is deleted is removed. Inserts need nothing,
 * because inserted rows/columns between the corners are inside the range automatically.
 */
export function adjustSheetObjectsForStructure(doc:Y.Doc,sheetId:string,sheet:Y.Array<Y.Array<string>>,axis:'row'|'column',index:number,deleting:boolean){
 if(!deleting||!doc.share.has('sheetObjects'))return;const map=doc.getMap<SheetObject>('sheetObjects'),rows=sheet.toArray(),columns=sheetColumnIds(sheet),key=(i:number)=>axis==='row'?sheetRowIdentity(rows[i]):columns[i];
 for(const [id,object]of sheetObjects(doc,sheetId)){
  const range=resolveSheetRange(sheet,object.range);if(!range){map.delete(id);continue;}
  const low=axis==='row'?range.top:range.left,high=axis==='row'?range.bottom:range.right;let next={...object.range};
  if(index<low||index>high){/* outside the range */}else if(low===high){map.delete(id);continue;}
  else if(index===low)next=axis==='row'?{...next,top:key(low+1),bottom:key(high)}:{...next,left:key(low+1),right:key(high)};
  else if(index===high)next=axis==='row'?{...next,top:key(low),bottom:key(high-1)}:{...next,left:key(low),right:key(high-1)};
  let value:SheetObject={...object,range:next};
  if(value.kind==='chart'&&(axis==='row'?value.anchor.row===key(index):value.anchor.column===key(index))){
   const at=index+1<(axis==='row'?rows.length:columns.length)?index+1:index-1;if(at<0){map.delete(id);continue;}value={...value,anchor:axis==='row'?{...value.anchor,row:key(at)}:{...value.anchor,column:key(at)}};
  }
  if(JSON.stringify(value)!==JSON.stringify(object))map.set(id,value);
 }
}
export function clearSheetObjects(doc:Y.Doc,sheetId:string){if(!doc.share.has('sheetObjects'))return;const map=doc.getMap('sheetObjects');for(const [id]of sheetObjects(doc,sheetId))map.delete(id);}

// ---- conditional formatting --------------------------------------------------------------------------------------
export type ConditionalCell={fill?:string;color?:string;bold?:true;bar?:{start:number;end:number;color:string}};
const numberOf=(text:string)=>{const trimmed=text.trim();if(!trimmed)return;const value=Number(trimmed);return Number.isFinite(value)?value:undefined;};
const mix=(a:string,b:string,t:number)=>[0,2,4].map(i=>Math.round(parseInt(a.slice(i,i+2),16)*(1-t)+parseInt(b.slice(i,i+2),16)*t).toString(16).padStart(2,'0')).join('').toUpperCase();
/**
 * Styles from the sheet's rules, evaluated on displayed values (formula results). Rules apply in priority order
 * (highest first); each property is taken from the first rule that sets it, like Excel's "first rule wins".
 */
export function evaluateConditionalFormats(rules:SheetConditionalFormat[],sheet:Y.Array<Y.Array<string>>,value:(row:number,column:number)=>string){
 const result=new Map<string,ConditionalCell>(),ordered=[...rules].sort((a,b)=>b.priority-a.priority);
 const assign=(key:string,style:ConditionalCell)=>{const cell=result.get(key)??{};for(const [name,v]of Object.entries(style))if(!(name in cell))(cell as Record<string,unknown>)[name]=v;result.set(key,cell);};
 for(const {range:keys,rule}of ordered){
  const range=resolveSheetRange(sheet,keys);if(!range)continue;const cells:{key:string;text:string;number?:number}[]=[];
  for(let r=range.top;r<=range.bottom&&cells.length<200000;r++)for(let c=range.left;c<=range.right;c++){const text=value(r,c);cells.push({key:r+':'+c,text,number:numberOf(text)});}
  const numbers=cells.flatMap(cell=>cell.number===undefined?[]:[cell.number]);
  if(rule.type==='colorScale'){
   if(!numbers.length)continue;const sorted=[...numbers].sort((a,b)=>a-b),min=sorted[0],max=sorted[sorted.length-1],mid=sorted.length%2?sorted[sorted.length>>1]:(sorted[sorted.length/2-1]+sorted[sorted.length/2])/2;
   for(const cell of cells){if(cell.number===undefined)continue;const [low,a,b]=rule.colors;let fill:string;
    if(max===min)fill=b??a;else if(b===undefined)fill=mix(low,a,(cell.number-min)/(max-min));else fill=cell.number<=mid?mix(low,a,mid===min?1:(cell.number-min)/(mid-min)):mix(a,b,max===mid?0:(cell.number-mid)/(max-mid));
    assign(cell.key,{fill});}
   continue;
  }
  if(rule.type==='dataBar'){
   if(!numbers.length)continue;const low=Math.min(0,...numbers),high=Math.max(0,...numbers),span=high-low||1,zero=(0-low)/span;
   for(const cell of cells)if(cell.number!==undefined){const at=(cell.number-low)/span;assign(cell.key,{bar:{start:Math.min(zero,at),end:Math.max(zero,at),color:rule.color}});}
   continue;
  }
  const test=conditionalTest(rule,cells,numbers);
  for(const cell of cells)if(test(cell))assign(cell.key,rule.style);
 }
 return result;
}
function conditionalTest(rule:Exclude<ConditionalRule,{type:'colorScale'|'dataBar'}>,cells:{text:string;number?:number}[],numbers:number[]):(cell:{text:string;number?:number})=>boolean{
 const lower=(text:string)=>text.toLowerCase();
 switch(rule.type){
  case 'greaterThan':case 'lessThan':case 'greaterOrEqual':case 'lessOrEqual':case 'equal':case 'notEqual':{
   const target=numberOf(rule.value),compare=(cell:{text:string;number?:number})=>target!==undefined&&cell.number!==undefined?cell.number-target:target===undefined&&cell.text!==''?(lower(cell.text)<lower(rule.value)?-1:lower(cell.text)>lower(rule.value)?1:0):undefined;
   return cell=>{const order=compare(cell);if(order===undefined)return rule.type==='notEqual'&&cell.text!==rule.value;
    return rule.type==='greaterThan'?order>0:rule.type==='lessThan'?order<0:rule.type==='greaterOrEqual'?order>=0:rule.type==='lessOrEqual'?order<=0:rule.type==='equal'?order===0:order!==0;};
  }
  case 'between':case 'notBetween':{const a=numberOf(rule.value),b=numberOf(rule.value2);if(a===undefined||b===undefined)return ()=>false;const low=Math.min(a,b),high=Math.max(a,b);return cell=>cell.number!==undefined&&((cell.number>=low&&cell.number<=high)===(rule.type==='between'));}
  case 'textContains':return cell=>lower(cell.text).includes(lower(rule.value));
  case 'textNotContains':return cell=>!lower(cell.text).includes(lower(rule.value));
  case 'beginsWith':return cell=>lower(cell.text).startsWith(lower(rule.value));
  case 'endsWith':return cell=>lower(cell.text).endsWith(lower(rule.value));
  case 'blank':return cell=>cell.text.trim()==='';
  case 'notBlank':return cell=>cell.text.trim()!=='';
  case 'duplicate':case 'unique':{const counts=new Map<string,number>();for(const cell of cells)if(cell.text!==''){const key=cell.number!==undefined?'n'+cell.number:'t'+lower(cell.text);counts.set(key,(counts.get(key)??0)+1);}
   return cell=>{if(cell.text==='')return false;const n=counts.get(cell.number!==undefined?'n'+cell.number:'t'+lower(cell.text))??0;return rule.type==='duplicate'?n>1:n===1;};}
  case 'top':case 'bottom':{if(!numbers.length)return ()=>false;const sorted=[...numbers].sort((a,b)=>rule.type==='top'?b-a:a-b),count=rule.percent?Math.max(1,Math.floor(numbers.length*rule.count/100)):Math.min(rule.count,numbers.length),edge=sorted[count-1];
   return cell=>cell.number!==undefined&&(rule.type==='top'?cell.number>=edge:cell.number<=edge);}
  case 'aboveAverage':case 'belowAverage':{if(!numbers.length)return ()=>false;const mean=numbers.reduce((a,b)=>a+b,0)/numbers.length;return cell=>cell.number!==undefined&&(rule.type==='aboveAverage'?cell.number>mean:cell.number<mean);}
 }
}

// ---- charts -------------------------------------------------------------------------------------------------------
export type ChartSeries={name:string;values:(number|null)[]};
/**
 * Chart data from the range: the first row and first column are headers when they hold text, as in Excel's quick
 * charts. Series run down columns unless `seriesInRows`. Non-numeric values are gaps.
 */
/** `value` is the raw computed value (numbers); `label` is the displayed text (formatted dates, headers). */
export function chartData(range:ResolvedRange,value:(row:number,column:number)=>string,seriesInRows=false,label=value){
 const height=range.bottom-range.top+1,width=range.right-range.left+1,at=(r:number,c:number)=>value(range.top+r,range.left+c),text=(r:number,c:number)=>label(range.top+r,range.left+c);
 const isText=(text:string)=>text.trim()!==''&&numberOf(text)===undefined;
 // Header row: text in the first row (beyond the label column). Header column: text in the first column below it.
 const headerRow=height>1&&Array.from({length:width},(_,c)=>at(0,c)).some((text,c)=>(width===1||c>0)&&isText(text));
 const headerColumn=width>1&&Array.from({length:height},(_,r)=>at(r,0)).some((text,r)=>(r>0||!headerRow)&&isText(text));
 const r0=headerRow?1:0,c0=headerColumn?1:0,rows=Math.min(height-r0,SHEET_OBJECT_LIMITS.points),columns=Math.min(width-c0,SHEET_OBJECT_LIMITS.points);
 const category=(i:number)=>seriesInRows?(headerRow?text(0,c0+i):String(i+1)):(headerColumn?text(r0+i,0):String(i+1));
 const seriesName=(i:number)=>seriesInRows?(headerColumn?text(r0+i,0):'系列'+(i+1)):(headerRow?text(0,c0+i):'系列'+(i+1));
 const count=Math.min(seriesInRows?rows:columns,SHEET_OBJECT_LIMITS.series),points=seriesInRows?columns:rows;
 const series:ChartSeries[]=Array.from({length:Math.max(0,count)},(_,s)=>({name:seriesName(s)||'系列'+(s+1),values:Array.from({length:Math.max(0,points)},(_,p)=>numberOf(seriesInRows?at(r0+s,c0+p):at(r0+p,c0+s))??null)}));
 return {categories:Array.from({length:Math.max(0,points)},(_,i)=>category(i)),series,headerRow,headerColumn};
}

// ---- data validation ------------------------------------------------------------------------------------------------
export function sheetValidations(doc:Y.Doc,sheetId:string){return sheetObjects(doc,sheetId).flatMap(([id,object])=>object.kind==='validation'?[[id,object as SheetValidation] as const]:[]);}
/** The validation covering a cell (the first stored one wins when ranges overlap). */
export function sheetValidationAt(doc:Y.Doc,sheet:Y.Array<Y.Array<string>>,sheetId:string,row:number,column:number){
 for(const [id,validation]of sheetValidations(doc,sheetId)){const range=resolveSheetRange(sheet,validation.range);if(range&&row>=range.top&&row<=range.bottom&&column>=range.left&&column<=range.right)return {id,validation,range};}
}
/** Whether a typed value passes the rule; blank cells always pass (Excel's “忽略空值”). */
export function validationAccepts(rule:SheetValidationRule,value:string){
 const text=value.trim();if(!text)return true;
 if(rule.type==='list')return rule.items.some(item=>item.toLowerCase()===text.toLowerCase());
 const amount=rule.type==='textLength'?[...value].length:Number(text);if(!Number.isFinite(amount)||rule.type==='whole'&&!Number.isInteger(amount))return false;
 const a=Number(rule.value),b=Number(rule.value2),low=Math.min(a,b),high=Math.max(a,b);
 switch(rule.operator){case 'between':return amount>=low&&amount<=high;case 'notBetween':return amount<low||amount>high;case 'equal':return amount===a;case 'notEqual':return amount!==a;case 'greaterThan':return amount>a;case 'lessThan':return amount<a;case 'greaterOrEqual':return amount>=a;case 'lessOrEqual':return amount<=a;}
}
const OPERATOR_TEXT:Record<SheetValidationOperator,string>={between:'介于',notBetween:'不介于',equal:'等于',notEqual:'不等于',greaterThan:'大于',lessThan:'小于',greaterOrEqual:'大于或等于',lessOrEqual:'小于或等于'};
/** Message shown when a value is rejected; a custom message wins. */
export function validationMessage(validation:SheetValidation){
 if(validation.message)return validation.message;const rule=validation.rule;
 if(rule.type==='list')return '输入值必须是下拉列表中的一项：'+rule.items.join('、');
 const what={whole:'整数',decimal:'数字',textLength:'文本长度'}[rule.type],range=rule.operator==='between'||rule.operator==='notBetween'?`${rule.value} 与 ${rule.value2}`:rule.value;
 return (rule.type==='textLength'?'文本长度必须':'输入值必须是'+what+'，且')+OPERATOR_TEXT[rule.operator]+' '+range;
}
