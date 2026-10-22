/** Function library shown in 公式 → 函数库 and the formula bar, grouped like Excel; every entry is implemented. */
export type SheetFunctionEntry={name:string;label:string;syntax:string};
export const SHEET_FUNCTION_GROUPS:{group:string;functions:SheetFunctionEntry[]}[]=[
 {group:'数学和统计',functions:[
  {name:'SUM',label:'求和',syntax:'SUM(数值1, [数值2], …)'},{name:'AVERAGE',label:'平均值',syntax:'AVERAGE(数值1, [数值2], …)'},
  {name:'COUNT',label:'数字个数',syntax:'COUNT(值1, [值2], …)'},{name:'COUNTA',label:'非空个数',syntax:'COUNTA(值1, [值2], …)'},{name:'COUNTBLANK',label:'空单元格个数',syntax:'COUNTBLANK(区域)'},
  {name:'MAX',label:'最大值',syntax:'MAX(数值1, [数值2], …)'},{name:'MIN',label:'最小值',syntax:'MIN(数值1, [数值2], …)'},{name:'MEDIAN',label:'中位数',syntax:'MEDIAN(数值1, [数值2], …)'},
  {name:'LARGE',label:'第 k 大',syntax:'LARGE(区域, k)'},{name:'SMALL',label:'第 k 小',syntax:'SMALL(区域, k)'},
  {name:'PRODUCT',label:'乘积',syntax:'PRODUCT(数值1, [数值2], …)'},{name:'SUMPRODUCT',label:'乘积之和',syntax:'SUMPRODUCT(区域1, [区域2], …)'},
  {name:'ROUND',label:'四舍五入',syntax:'ROUND(数值, 小数位数)'},{name:'ROUNDUP',label:'向上舍入',syntax:'ROUNDUP(数值, 小数位数)'},{name:'ROUNDDOWN',label:'向下舍入',syntax:'ROUNDDOWN(数值, 小数位数)'},
  {name:'INT',label:'向下取整',syntax:'INT(数值)'},{name:'CEILING',label:'按倍数向上舍入',syntax:'CEILING(数值, [倍数])'},{name:'FLOOR',label:'按倍数向下舍入',syntax:'FLOOR(数值, [倍数])'},
  {name:'MOD',label:'余数',syntax:'MOD(被除数, 除数)'},{name:'ABS',label:'绝对值',syntax:'ABS(数值)'},{name:'SIGN',label:'符号',syntax:'SIGN(数值)'},
  {name:'POWER',label:'乘幂',syntax:'POWER(底数, 指数)'},{name:'SQRT',label:'平方根',syntax:'SQRT(数值)'},{name:'PI',label:'圆周率',syntax:'PI()'},
 ]},
 {group:'条件统计',functions:[
  {name:'SUMIF',label:'条件求和',syntax:'SUMIF(条件区域, 条件, [求和区域])'},{name:'SUMIFS',label:'多条件求和',syntax:'SUMIFS(求和区域, 条件区域1, 条件1, …)'},
  {name:'COUNTIF',label:'条件计数',syntax:'COUNTIF(区域, 条件)'},{name:'COUNTIFS',label:'多条件计数',syntax:'COUNTIFS(条件区域1, 条件1, …)'},
  {name:'AVERAGEIF',label:'条件平均',syntax:'AVERAGEIF(条件区域, 条件, [平均区域])'},{name:'AVERAGEIFS',label:'多条件平均',syntax:'AVERAGEIFS(平均区域, 条件区域1, 条件1, …)'},
  {name:'MAXIFS',label:'条件最大值',syntax:'MAXIFS(最大值区域, 条件区域1, 条件1, …)'},{name:'MINIFS',label:'条件最小值',syntax:'MINIFS(最小值区域, 条件区域1, 条件1, …)'},
 ]},
 {group:'逻辑',functions:[
  {name:'IF',label:'条件判断',syntax:'IF(条件, 为真时, [为假时])'},{name:'IFS',label:'多条件判断',syntax:'IFS(条件1, 值1, [条件2, 值2], …)'},
  {name:'IFERROR',label:'出错时返回',syntax:'IFERROR(值, 出错时的值)'},{name:'IFNA',label:'#N/A 时返回',syntax:'IFNA(值, #N/A 时的值)'},
  {name:'AND',label:'全部为真',syntax:'AND(条件1, [条件2], …)'},{name:'OR',label:'任一为真',syntax:'OR(条件1, [条件2], …)'},{name:'XOR',label:'奇数个为真',syntax:'XOR(条件1, [条件2], …)'},
  {name:'NOT',label:'取反',syntax:'NOT(条件)'},{name:'SWITCH',label:'按值选择',syntax:'SWITCH(表达式, 值1, 结果1, …, [默认值])'},
  {name:'TRUE',label:'逻辑真',syntax:'TRUE()'},{name:'FALSE',label:'逻辑假',syntax:'FALSE()'},
 ]},
 {group:'查找与引用',functions:[
  {name:'VLOOKUP',label:'按列查找',syntax:'VLOOKUP(查找值, 表格区域, 列序号, [近似匹配])'},{name:'HLOOKUP',label:'按行查找',syntax:'HLOOKUP(查找值, 表格区域, 行序号, [近似匹配])'},
  {name:'XLOOKUP',label:'查找并返回',syntax:'XLOOKUP(查找值, 查找区域, 返回区域, [未找到时], [匹配方式], [搜索方式])'},
  {name:'MATCH',label:'查找位置',syntax:'MATCH(查找值, 查找区域, [匹配类型])'},{name:'INDEX',label:'按位置取值',syntax:'INDEX(区域, 行号, [列号])'},
 ]},
 {group:'文本',functions:[
  {name:'CONCAT',label:'连接文本',syntax:'CONCAT(文本1, [文本2], …)'},{name:'CONCATENATE',label:'连接文本（旧版）',syntax:'CONCATENATE(文本1, [文本2], …)'},{name:'TEXTJOIN',label:'用分隔符连接',syntax:'TEXTJOIN(分隔符, 忽略空值, 文本1, …)'},
  {name:'LEFT',label:'左侧字符',syntax:'LEFT(文本, [字符数])'},{name:'RIGHT',label:'右侧字符',syntax:'RIGHT(文本, [字符数])'},{name:'MID',label:'中间字符',syntax:'MID(文本, 开始位置, 字符数)'},
  {name:'LEN',label:'字符数',syntax:'LEN(文本)'},{name:'FIND',label:'查找（区分大小写）',syntax:'FIND(查找文本, 文本, [开始位置])'},{name:'SEARCH',label:'查找（支持通配符）',syntax:'SEARCH(查找文本, 文本, [开始位置])'},
  {name:'SUBSTITUTE',label:'替换文本',syntax:'SUBSTITUTE(文本, 旧文本, 新文本, [第几处])'},{name:'REPLACE',label:'按位置替换',syntax:'REPLACE(文本, 开始位置, 字符数, 新文本)'},
  {name:'TRIM',label:'清除多余空格',syntax:'TRIM(文本)'},{name:'UPPER',label:'转大写',syntax:'UPPER(文本)'},{name:'LOWER',label:'转小写',syntax:'LOWER(文本)'},{name:'PROPER',label:'首字母大写',syntax:'PROPER(文本)'},
  {name:'TEXT',label:'按格式显示',syntax:'TEXT(值, 格式)'},{name:'VALUE',label:'文本转数值',syntax:'VALUE(文本)'},{name:'EXACT',label:'完全相同',syntax:'EXACT(文本1, 文本2)'},{name:'REPT',label:'重复文本',syntax:'REPT(文本, 次数)'},
 ]},
 {group:'日期和时间',functions:[
  {name:'TODAY',label:'今天',syntax:'TODAY()'},{name:'NOW',label:'当前日期时间',syntax:'NOW()'},{name:'DATE',label:'构造日期',syntax:'DATE(年, 月, 日)'},{name:'DATEVALUE',label:'文本转日期',syntax:'DATEVALUE(日期文本)'},
  {name:'YEAR',label:'年',syntax:'YEAR(日期)'},{name:'MONTH',label:'月',syntax:'MONTH(日期)'},{name:'DAY',label:'日',syntax:'DAY(日期)'},
  {name:'HOUR',label:'时',syntax:'HOUR(时间)'},{name:'MINUTE',label:'分',syntax:'MINUTE(时间)'},{name:'SECOND',label:'秒',syntax:'SECOND(时间)'},
  {name:'WEEKDAY',label:'星期几',syntax:'WEEKDAY(日期, [类型])'},{name:'DAYS',label:'相差天数',syntax:'DAYS(结束日期, 开始日期)'},{name:'DATEDIF',label:'日期间隔',syntax:'DATEDIF(开始日期, 结束日期, "Y"|"M"|"D"|"YM"|"MD"|"YD")'},
  {name:'EDATE',label:'几个月后的日期',syntax:'EDATE(开始日期, 月数)'},{name:'EOMONTH',label:'月末日期',syntax:'EOMONTH(开始日期, 月数)'},
 ]},
 {group:'信息',functions:[
  {name:'ISBLANK',label:'是否为空',syntax:'ISBLANK(值)'},{name:'ISNUMBER',label:'是否为数字',syntax:'ISNUMBER(值)'},{name:'ISTEXT',label:'是否为文本',syntax:'ISTEXT(值)'},
  {name:'ISLOGICAL',label:'是否为逻辑值',syntax:'ISLOGICAL(值)'},{name:'ISERROR',label:'是否为错误',syntax:'ISERROR(值)'},{name:'ISNA',label:'是否为 #N/A',syntax:'ISNA(值)'},
 ]},
];
const index=new Map(SHEET_FUNCTION_GROUPS.flatMap(group=>group.functions.map(entry=>[entry.name,entry] as const)));
export const sheetFunctionEntry=(name:string)=>index.get(name.toUpperCase());
/** The innermost function call the caret is in (`=IF(A1>0,ROUND(|` → ROUND) and which argument, for syntax hints. */
export function activeSheetFunction(formula:string,caret:number){
 const before=formula.slice(0,caret);if(!before.trimStart().startsWith('='))return;
 const stack:{name:string;argument:number}[]=[];let quote:string|undefined;
 for(let i=0;i<before.length;i++){const char=before[i];
  if(quote){if(char===quote)quote=undefined;continue;}if(char==='"'||char==="'"){quote=char;continue;}
  if(char==='('){const name=/([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(before.slice(0,i))?.[1];stack.push({name:(name??'').toUpperCase().replace(/^_XLFN\./,''),argument:0});}
  else if(char===')')stack.pop();else if(char===','&&stack.length)stack[stack.length-1].argument++;
 }
 for(let i=stack.length-1;i>=0;i--){const entry=sheetFunctionEntry(stack[i].name);if(entry)return {entry,argument:stack[i].argument};}
}
