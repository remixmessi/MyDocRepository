import {FormulaError,fail,finite,numeric,type Ast,type ErrorCode,type FormulaContext,type FormulaFunction,type FormulaMatrix,type Scalar} from './spreadsheet-formula-core';
import {excelSerialParts,excelSerialFromParts,formatSheetValue,parseSheetDateInput} from './spreadsheet-format';

/**
 * Excel-compatible function library (common math/statistics, conditional aggregation, logical, lookup, text and date
 * functions). Semantics follow Excel where it matters for results; deliberate differences are noted inline.
 */
export const TEXT_LIMIT=32767;
const PATTERN_LIMIT=255;
const RECOVERABLE=new Set<ErrorCode>(['#VALUE!','#DIV/0!','#REF!','#NAME?','#NUM!','#N/A']);

// ---- coercion ---------------------------------------------------------------------------------------------------
/** Number → text like Excel's General format: 15 significant digits, so 0.1+0.2 joins as "0.3". */
export function numberText(value:number){if(Number.isInteger(value)&&Math.abs(value)<1e15)return String(value);const text=Number(value.toPrecision(15)).toString();return text==='-0'?'0':text;}
export function toText(value:Scalar):string{return value===null?'':typeof value==='boolean'?value?'TRUE':'FALSE':typeof value==='number'?numberText(value):value;}
function truth(value:Scalar){if(typeof value==='string'){const upper=value.toUpperCase();if(upper==='TRUE')return true;if(upper==='FALSE')return false;return fail('#VALUE!');}return Boolean(value);}
const integer=(value:Scalar)=>Math.trunc(numeric(value));
const chars=(value:string)=>Array.from(value);
function limited(text:string,ctx:FormulaContext){ctx.work(text.length);if(text.length>TEXT_LIMIT)fail('#VALUE!');return text;}
/** Numbers from arguments: range cells count only when numeric; literal arguments are coerced (SUM(1,"2")=3). */
function numbers(args:Ast[],ctx:FormulaContext){
 const out:number[]=[];
 for(const arg of args){const values=ctx.values(arg);for(const value of values){if(typeof value==='number')out.push(value);else if(arg.kind==='literal'&&value!==null)out.push(numeric(value));else if(arg.kind!=='ref'&&arg.kind!=='range'&&typeof value==='boolean')out.push(value?1:0);}}
 return out;
}
const sum=(list:number[])=>finite(list.reduce((a,b)=>a+b,0));

// ---- comparison -------------------------------------------------------------------------------------------------
const rank=(value:Scalar)=>typeof value==='number'||value===null?0:typeof value==='string'?1:2;
/** Excel ordering across types: numbers < text < logical; text compares case-insensitively and locale-independently. */
export function compareScalars(a:Scalar,b:Scalar){
 const ra=rank(a),rb=rank(b);if(ra!==rb)return ra-rb;
 if(ra===0)return (a as number??0)-(b as number??0);
 if(ra===1){const x=(a as string).toLowerCase(),y=(b as string).toLowerCase();return x<y?-1:x>y?1:0;}
 return Number(a)-Number(b);
}
function sameValue(a:Scalar,b:Scalar){if((a===null||a==='')&&(b===null||b===''))return true;if(a===null||b===null)return false;return typeof a===typeof b&&(typeof a==='string'?a.toLowerCase()===(b as string).toLowerCase():a===b);}
/** Excel wildcards (* ? and ~ escape), matched linearly with backtracking to the last star; work is charged per step. */
export function wildcardMatch(text:string,pattern:string,ctx:FormulaContext){
 if(pattern.length>PATTERN_LIMIT)fail('#VALUE!');const t=chars(text.toLowerCase()),tokens:{char?:string;any?:boolean;one?:boolean}[]=[];
 const p=chars(pattern.toLowerCase());for(let i=0;i<p.length;i++){if(p[i]==='~'&&i+1<p.length){tokens.push({char:p[++i]});continue;}tokens.push(p[i]==='*'?{any:true}:p[i]==='?'?{one:true}:{char:p[i]});}
 let ti=0,pi=0,star=-1,mark=0;
 while(ti<t.length){ctx.work(1);const token=tokens[pi];
  if(token&&(token.one||token.char===t[ti])){ti++;pi++;}else if(token?.any){star=pi++;mark=ti;}else if(star>=0){pi=star+1;ti=++mark;}else return false;}
 while(tokens[pi]?.any)pi++;return pi===tokens.length;
}
const hasWildcard=(text:string)=>/(^|[^~])[*?]/.test(text);
/** SUMIF/COUNTIF criteria: ">5", "<>", "=x", wildcard text, numbers; numeric criteria also match numeric text. */
export function criterion(value:Scalar,ctx:FormulaContext):(cell:Scalar)=>boolean{
 if(typeof value!=='string'){const target=value===null?0:value;return cell=>typeof target==='number'?(typeof cell==='number'&&cell===target)||(typeof cell==='string'&&cell.trim()!==''&&Number(cell)===target):typeof cell==='boolean'&&cell===target;}
 const match=/^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(value)!,op=match[1]??'',operand=match[2],number=operand.trim()!==''&&Number.isFinite(Number(operand))?Number(operand):undefined;
 const blank=(cell:Scalar)=>cell===null||cell==='';
 if(operand===''){if(op==='<>')return cell=>!blank(cell);if(op===''||op==='=')return blank;return ()=>false;}
 if(['<','>','<=','>='].includes(op)){
  const test=(n:number)=>op==='<'?n<0:op==='>'?n>0:op==='<='?n<=0:n>=0;
  if(number!==undefined)return cell=>typeof cell==='number'&&test(cell-number);
  return cell=>typeof cell==='string'&&cell!==''&&test(compareScalars(cell,operand));
 }
 const logical=operand.toUpperCase()==='TRUE'?true:operand.toUpperCase()==='FALSE'?false:undefined;
 const equal=(cell:Scalar)=>{
  if(number!==undefined&&(typeof cell==='number'||typeof cell==='string'&&cell.trim()!==''&&Number.isFinite(Number(cell))))return Number(cell)===number;
  if(logical!==undefined&&typeof cell==='boolean')return cell===logical;
  if(typeof cell!=='string')return false;return hasWildcard(operand)?wildcardMatch(cell,operand,ctx):cell.toLowerCase()===operand.replace(/~([*?~])/g,'$1').toLowerCase();
 };
 return op==='<>'?cell=>!equal(cell):equal;
}
/** Cell value for matching; an error inside a criteria range simply does not match (as in Excel). */
function safe(matrix:FormulaMatrix,row:number,column:number):Scalar|FormulaError{
 try{return matrix.get(row,column);}catch(error){if(error instanceof FormulaError&&RECOVERABLE.has(error.code))return error;throw error;}
}
/** Positions where every (range, criterion) pair matches; all criteria ranges must share the first range's shape. */
function matchedCells(args:Ast[],ctx:FormulaContext){
 if(args.length%2)fail('#VALUE!');const pairs:{matrix:FormulaMatrix;test:(cell:Scalar)=>boolean}[]=[];
 for(let i=0;i<args.length;i+=2){const matrix=ctx.matrix(args[i]);if(pairs.length&&(matrix.rows!==pairs[0].matrix.rows||matrix.columns!==pairs[0].matrix.columns))fail('#VALUE!');pairs.push({matrix,test:criterion(ctx.value(args[i+1]),ctx)});}
 const first=pairs[0].matrix,out:[number,number][]=[];
 for(let r=0;r<first.rows;r++)for(let c=0;c<first.columns;c++){ctx.tick();if(pairs.every(({matrix,test})=>{const cell=safe(matrix,r,c);return !(cell instanceof FormulaError)&&test(cell);}))out.push([r,c]);}
 return {cells:out,rows:first.rows,columns:first.columns};
}
/** Numeric values of `target` at matched positions; errors in matched cells propagate. */
function matchedNumbers(target:FormulaMatrix,cells:[number,number][]){const out:number[]=[];for(const [r,c]of cells){const value=target.get(r,c);if(typeof value==='number')out.push(value);}return out;}

// ---- dates ------------------------------------------------------------------------------------------------------
const DAY_MS=86400000,EPOCH=Date.UTC(1899,11,30);
/** Date serial from a number, a date text ("2026-10-21") or a numeric text. */
export function serialOf(value:Scalar){
 if(typeof value==='boolean')return fail('#VALUE!');if(typeof value==='number')return value;if(value===null||value==='')return 0;
 const parsed=parseSheetDateInput(value);if(parsed!==undefined)return Number(parsed);const number=Number(value.trim());return value.trim()&&Number.isFinite(number)?number:fail('#VALUE!');
}
function parts(value:Scalar){const serial=serialOf(value);if(serial<0)fail('#NUM!');return excelSerialParts(serial)??fail('#NUM!');}
/** DATE(y,m,d): months and days overflow into the next unit; years 0–1899 add 1900, as in Excel. */
function dateSerial(year:number,month:number,day:number){
 let y=Math.trunc(year);if(y>=0&&y<1900)y+=1900;if(y<1900||y>9999)fail('#NUM!');
 const utc=Date.UTC(y,Math.trunc(month)-1,1)+(Math.trunc(day)-1)*DAY_MS,date=new Date(utc);
 return excelSerialFromParts(date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate())??fail('#NUM!');
}
const lastDay=(year:number,month:number)=>new Date(Date.UTC(year,month,0)).getUTCDate();
function addMonths(value:Scalar,months:number,endOfMonth:boolean){
 const p=parts(value),total=p.year*12+(p.month-1)+Math.trunc(months),year=Math.floor(total/12),month=total%12+1;
 if(year<1900||year>9999)fail('#NUM!');return excelSerialFromParts(year,month,endOfMonth?lastDay(year,month):Math.min(p.day,lastDay(year,month)))??fail('#NUM!');
}
/** Local wall-clock time, matching what each user sees (Excel's TODAY/NOW are local too). */
function nowSerial(withTime:boolean){
 const now=new Date(),utc=Date.UTC(now.getFullYear(),now.getMonth(),now.getDate()),days=Math.round((utc-EPOCH)/DAY_MS)-(utc<Date.UTC(1900,2,1)?1:0);
 return withTime?days+(now.getHours()*3600+now.getMinutes()*60+now.getSeconds())/86400:days;
}
function datedif(start:number,end:number,unit:string){
 if(start>end)fail('#NUM!');const a=parts(start),b=parts(end);
 const months=(b.year-a.year)*12+(b.month-a.month)-(b.day<a.day?1:0);
 switch(unit.toUpperCase()){
  case 'D':return Math.floor(end)-Math.floor(start);
  case 'M':return months;
  case 'Y':return Math.floor(months/12);
  case 'YM':return months%12;
  case 'MD':return b.day>=a.day?b.day-a.day:b.day+lastDay(b.month===1?b.year-1:b.year,b.month===1?12:b.month-1)-a.day;
  case 'YD':{let anchor=excelSerialFromParts(b.year,a.month,Math.min(a.day,lastDay(b.year,a.month)))!;if(anchor>Math.floor(end))anchor=excelSerialFromParts(b.year-1,a.month,Math.min(a.day,lastDay(b.year-1,a.month)))!;return Math.floor(end)-anchor;}
  default:return fail('#NUM!');
 }
}

// ---- rounding -----------------------------------------------------------------------------------------------------
/** Rounding on the decimal representation (15 significant digits) so ROUNDUP(1.1,1) is 1.1, not 1.2. */
function roundDigits(value:number,digits:number,mode:'round'|'up'|'down'){
 if(Math.abs(digits)>308)fail('#NUM!');const factor=10**digits,scaled=Number((Math.abs(value)*factor).toPrecision(15));
 // Scale back with an exact power of ten: 12*100 is exact, 12/0.01 is not.
 const rounded=mode==='round'?Math.round(scaled):mode==='up'?Math.ceil(scaled):Math.floor(scaled);return finite(Math.sign(value)*(digits>=0?rounded/10**digits:rounded*10**-digits));
}
function multiple(value:number,step:number,up:boolean){
 // CEILING(-2.5,-2)=-4 and FLOOR(-2.5,-2)=-2 fall out of ceil/floor on value/step; a positive value with a negative step is #NUM!.
 if(step===0)return 0;if(value>0&&step<0)fail('#NUM!');const ratio=Number((value/step).toPrecision(15));
 return finite(Number(((up?Math.ceil(ratio):Math.floor(ratio))*step).toPrecision(15)));
}

// ---- lookup -----------------------------------------------------------------------------------------------------
type Line={length:number;get(index:number):Scalar|FormulaError};
const lineOf=(matrix:FormulaMatrix,axis:'row'|'column',index=0):Line=>axis==='column'?{length:matrix.rows,get:i=>safe(matrix,i,index)}:{length:matrix.columns,get:i=>safe(matrix,index,i)};
/** Excel: looking up an empty cell is #N/A rather than matching blanks. */
function lookupKey(ctx:FormulaContext,node:Ast){const key=ctx.value(node);return key===null?fail('#N/A'):key;}
function exactIndex(line:Line,key:Scalar,ctx:FormulaContext,wildcards:boolean,reverse=false){
 const pattern=wildcards&&typeof key==='string'&&hasWildcard(key);
 for(let n=0;n<line.length;n++){const i=reverse?line.length-1-n:n;ctx.tick();const value=line.get(i);if(value instanceof FormulaError||value===null)continue;if(pattern?typeof value==='string'&&wildcardMatch(value,key as string,ctx):sameValue(value,key))return i;}
 return -1;
}
/** Sorted approximate match: last position not greater (ascending) or not smaller (descending) than the key. */
function sortedIndex(line:Line,key:Scalar,ctx:FormulaContext,descending=false){
 let found=-1;
 for(let i=0;i<line.length;i++){ctx.tick();const value=line.get(i);if(value instanceof FormulaError||value===null||rank(value)!==rank(key))continue;const order=compareScalars(value,key);if(descending?order>=0:order<=0)found=i;else break;}
 return found;
}
/** XLOOKUP next smaller (-1) / next larger (1) on unsorted data. */
function nearestIndex(line:Line,key:Scalar,ctx:FormulaContext,direction:-1|1,reverse:boolean){
 const exact=exactIndex(line,key,ctx,false,reverse);if(exact>=0)return exact;let best=-1,bestValue:Scalar=null;
 for(let n=0;n<line.length;n++){const i=reverse?line.length-1-n:n;ctx.tick();const value=line.get(i);if(value instanceof FormulaError||value===null||rank(value)!==rank(key))continue;
  const order=compareScalars(value,key);if(direction<0?order<0:order>0)if(best<0||(direction<0?compareScalars(value,bestValue)>0:compareScalars(value,bestValue)<0)){best=i;bestValue=value;}}
 return best;
}
function attempt(ctx:FormulaContext,node:Ast):Scalar|FormulaError{try{return ctx.value(node);}catch(error){if(error instanceof FormulaError&&RECOVERABLE.has(error.code))return error;throw error;}}
/** Every cell of an argument, keeping cell errors as values (COUNTA counts them). */
function cellsOf(node:Ast,ctx:FormulaContext):(Scalar|FormulaError)[]{
 if(node.kind!=='ref'&&node.kind!=='range')return [attempt(ctx,node)];const matrix=ctx.matrix(node),out:(Scalar|FormulaError)[]=[];
 for(let r=0;r<matrix.rows;r++)for(let c=0;c<matrix.columns;c++){ctx.tick();out.push(safe(matrix,r,c));}return out;
}
function logicals(args:Ast[],ctx:FormulaContext){
 const out:boolean[]=[];for(const arg of args){const range=arg.kind==='ref'||arg.kind==='range';for(const value of ctx.values(arg)){if(value===null||range&&typeof value==='string')continue;out.push(truth(value));}}
 return out.length?out:fail('#VALUE!');
}

// ---- function table -----------------------------------------------------------------------------------------------
// Excel allows 255 arguments; existing workbooks may hold more, which the token limit (2048) still bounds.
const MANY=1024;
const fn=(min:number,max:number,run:FormulaFunction['run']):FormulaFunction=>({min,max,run});
const text=(ctx:FormulaContext,node:Ast)=>toText(ctx.value(node));
/** An optional argument that was actually written (not absent and not left empty). */
const given=(node:Ast|undefined)=>node!==undefined&&!(node.kind==='literal'&&node.omitted);
const num=(ctx:FormulaContext,node:Ast)=>numeric(ctx.value(node));
const ifs=(aggregate:(list:number[])=>Scalar)=>fn(3,MANY,(args,ctx)=>{const {cells,rows,columns}=matchedCells(args.slice(1),ctx),target=ctx.matrix(args[0]);if(target.rows!==rows||target.columns!==columns)fail('#VALUE!');return aggregate(matchedNumbers(target,cells));});
const ifOne=(aggregate:(list:number[])=>Scalar)=>fn(2,3,(args,ctx)=>{const {cells,rows,columns}=matchedCells([args[0],args[1]],ctx);return aggregate(matchedNumbers(args[2]?ctx.matrix(args[2],rows,columns):ctx.matrix(args[0]),cells));});
const average=(list:number[])=>list.length?finite(sum(list)/list.length):fail('#DIV/0!');
const datePart=(key:'year'|'month'|'day'|'hour'|'minute'|'second')=>fn(1,1,(args,ctx)=>parts(ctx.value(args[0]))[key]);
const is=(test:(value:Scalar|FormulaError)=>boolean)=>fn(1,1,(args,ctx)=>test(attempt(ctx,args[0])));
const kth=(largest:boolean)=>fn(2,2,(args,ctx)=>{const list=numbers([args[0]],ctx).sort((a,b)=>largest?b-a:a-b),k=Math.trunc(num(ctx,args[1]));return k>=1&&k<=list.length?list[k-1]:fail('#NUM!');});
const round=(mode:'round'|'up'|'down')=>fn(2,2,(args,ctx)=>roundDigits(num(ctx,args[0]),Math.trunc(num(ctx,args[1])),mode));
const lookupTable=(axis:'row'|'column')=>fn(3,4,(args,ctx)=>{
 const key=lookupKey(ctx,args[0]),table=ctx.matrix(args[1]),index=integer(ctx.value(args[2])),approximate=args[3]===undefined?true:truth(ctx.value(args[3]));
 if(index<1)fail('#VALUE!');if(index>(axis==='column'?table.columns:table.rows))fail('#REF!');
 const line=lineOf(table,axis),at=approximate?sortedIndex(line,key,ctx):exactIndex(line,key,ctx,true);if(at<0)fail('#N/A');
 return axis==='column'?table.get(at,index-1):table.get(index-1,at);
});

export const FORMULA_FUNCTIONS:Record<string,FormulaFunction>={
 // Math and statistics
 SUM:fn(1,MANY,(args,ctx)=>sum(numbers(args,ctx))),
 AVERAGE:fn(1,MANY,(args,ctx)=>average(numbers(args,ctx))),
 COUNT:fn(1,MANY,(args,ctx)=>numbers(args,ctx).length),
 COUNTA:fn(1,MANY,(args,ctx)=>args.flatMap(arg=>cellsOf(arg,ctx)).filter(value=>value!==null).length),
 COUNTBLANK:fn(1,1,(args,ctx)=>cellsOf(args[0],ctx).filter(value=>value===null||value==='').length),
 MIN:fn(1,MANY,(args,ctx)=>{const list=numbers(args,ctx);return list.length?list.reduce((a,b)=>Math.min(a,b)):0;}),
 MAX:fn(1,MANY,(args,ctx)=>{const list=numbers(args,ctx);return list.length?list.reduce((a,b)=>Math.max(a,b)):0;}),
 MEDIAN:fn(1,MANY,(args,ctx)=>{const list=numbers(args,ctx).sort((a,b)=>a-b);if(!list.length)fail('#NUM!');const mid=list.length>>1;return list.length%2?list[mid]:(list[mid-1]+list[mid])/2;}),
 LARGE:kth(true),SMALL:kth(false),
 PRODUCT:fn(1,MANY,(args,ctx)=>{const list=numbers(args,ctx);return list.length?finite(list.reduce((a,b)=>a*b,1)):0;}),
 SUMPRODUCT:fn(1,30,(args,ctx)=>{const matrices=args.map(arg=>ctx.matrix(arg)),[first]=matrices;if(matrices.some(m=>m.rows!==first.rows||m.columns!==first.columns))fail('#VALUE!');let total=0;
  for(let r=0;r<first.rows;r++)for(let c=0;c<first.columns;c++){ctx.tick();total+=matrices.reduce((product,m)=>{const value=m.get(r,c);return product*(typeof value==='number'?value:0);},1);}return finite(total);}),
 ROUND:round('round'),ROUNDUP:round('up'),ROUNDDOWN:round('down'),
 INT:fn(1,1,(args,ctx)=>Math.floor(num(ctx,args[0]))),
 ABS:fn(1,1,(args,ctx)=>Math.abs(num(ctx,args[0]))),
 SIGN:fn(1,1,(args,ctx)=>Math.sign(num(ctx,args[0]))||0),
 MOD:fn(2,2,(args,ctx)=>{const a=num(ctx,args[0]),b=num(ctx,args[1]);if(b===0)fail('#DIV/0!');return finite(Number((a-b*Math.floor(a/b)).toPrecision(15)));}),
 POWER:fn(2,2,(args,ctx)=>{const a=num(ctx,args[0]),b=num(ctx,args[1]);if(a===0&&b<0)fail('#DIV/0!');return finite(a**b);}),
 SQRT:fn(1,1,(args,ctx)=>{const a=num(ctx,args[0]);return a<0?fail('#NUM!'):Math.sqrt(a);}),
 PI:fn(0,0,()=>Math.PI),
 CEILING:fn(1,2,(args,ctx)=>multiple(num(ctx,args[0]),args[1]?num(ctx,args[1]):1,true)),
 FLOOR:fn(1,2,(args,ctx)=>multiple(num(ctx,args[0]),args[1]?num(ctx,args[1]):1,false)),
 // Conditional aggregation
 SUMIF:ifOne(sum),AVERAGEIF:ifOne(average),
 COUNTIF:fn(2,2,(args,ctx)=>matchedCells(args,ctx).cells.length),
 COUNTIFS:fn(2,MANY,(args,ctx)=>matchedCells(args,ctx).cells.length),
 SUMIFS:ifs(sum),AVERAGEIFS:ifs(average),
 MAXIFS:ifs(list=>list.length?Math.max(...list):0),MINIFS:ifs(list=>list.length?Math.min(...list):0),
 // Logical
 IF:fn(2,3,(args,ctx)=>truth(ctx.value(args[0]))?ctx.value(args[1]):args[2]?ctx.value(args[2]):false),
 IFS:fn(2,MANY,(args,ctx)=>{if(args.length%2)fail('#VALUE!');for(let i=0;i<args.length;i+=2)if(truth(ctx.value(args[i])))return ctx.value(args[i+1]);return fail('#N/A');}),
 AND:fn(1,MANY,(args,ctx)=>logicals(args,ctx).every(Boolean)),
 OR:fn(1,MANY,(args,ctx)=>logicals(args,ctx).some(Boolean)),
 XOR:fn(1,MANY,(args,ctx)=>logicals(args,ctx).filter(Boolean).length%2===1),
 NOT:fn(1,1,(args,ctx)=>!truth(ctx.value(args[0]))),
 IFERROR:fn(2,2,(args,ctx)=>{const value=attempt(ctx,args[0]);return value instanceof FormulaError?ctx.value(args[1]):value;}),
 IFNA:fn(2,2,(args,ctx)=>{const value=attempt(ctx,args[0]);if(value instanceof FormulaError){if(value.code==='#N/A')return ctx.value(args[1]);throw value;}return value;}),
 SWITCH:fn(3,MANY,(args,ctx)=>{const target=ctx.value(args[0]);let i=1;for(;i+1<args.length;i+=2)if(sameValue(ctx.value(args[i]),target))return ctx.value(args[i+1]);return i<args.length?ctx.value(args[i]):fail('#N/A');}),
 TRUE:fn(0,0,()=>true),FALSE:fn(0,0,()=>false),
 ISBLANK:is(value=>value===null),ISNUMBER:is(value=>typeof value==='number'),ISTEXT:is(value=>typeof value==='string'),
 ISLOGICAL:is(value=>typeof value==='boolean'),ISERROR:is(value=>value instanceof FormulaError),ISNA:is(value=>value instanceof FormulaError&&value.code==='#N/A'),
 // Lookup and reference
 VLOOKUP:lookupTable('column'),HLOOKUP:lookupTable('row'),
 MATCH:fn(2,3,(args,ctx)=>{
  const key=lookupKey(ctx,args[0]),matrix=ctx.matrix(args[1]),type=args[2]===undefined?1:integer(ctx.value(args[2]));if(matrix.rows>1&&matrix.columns>1)fail('#N/A');
  const line=lineOf(matrix,matrix.rows===1?'row':'column'),at=type===0?exactIndex(line,key,ctx,true):type>0?sortedIndex(line,key,ctx):sortedIndex(line,key,ctx,true);return at<0?fail('#N/A'):at+1;}),
 INDEX:fn(2,3,(args,ctx)=>{
  const matrix=ctx.matrix(args[0]),a=integer(ctx.value(args[1])),b=args[2]===undefined?undefined:integer(ctx.value(args[2]));let row:number,column:number;
  // INDEX(range,0,…) would return a whole row/column, which a single cell cannot hold: #REF! here.
  if(b===undefined){if(matrix.rows===1){row=1;column=a;}else if(matrix.columns===1){row=a;column=1;}else return fail('#REF!');}else{row=a;column=b;}
  if(row<1||column<1||row>matrix.rows||column>matrix.columns)fail('#REF!');return matrix.get(row-1,column-1);}),
 XLOOKUP:fn(3,6,(args,ctx)=>{
  const key=lookupKey(ctx,args[0]),keys=ctx.matrix(args[1]),values=ctx.matrix(args[2]);if(keys.rows>1&&keys.columns>1)fail('#VALUE!');
  const vertical=keys.columns===1&&(keys.rows>1||values.columns===1),length=vertical?keys.rows:keys.columns;
  // The return range must be one row/column of the same length (a 2-D return range is not supported).
  if((vertical?values.rows:values.columns)!==length||(vertical?values.columns:values.rows)!==1)fail('#VALUE!');
  const mode=args[4]===undefined?0:integer(ctx.value(args[4])),search=args[5]===undefined?1:integer(ctx.value(args[5]));if(![0,-1,1,2].includes(mode)||![1,-1,2,-2].includes(search))fail('#VALUE!');
  const line=lineOf(keys,vertical?'column':'row'),reverse=search<0,at=mode===0?exactIndex(line,key,ctx,false,reverse):mode===2?exactIndex(line,key,ctx,true,reverse):nearestIndex(line,key,ctx,mode as -1|1,reverse);
  if(at<0)return given(args[3])?ctx.value(args[3]):fail('#N/A');return vertical?values.get(at,0):values.get(0,at);}),
 // Text
 LEN:fn(1,1,(args,ctx)=>chars(text(ctx,args[0])).length),
 LEFT:fn(1,2,(args,ctx)=>{const count=args[1]?integer(ctx.value(args[1])):1;if(count<0)fail('#VALUE!');return chars(text(ctx,args[0])).slice(0,count).join('');}),
 RIGHT:fn(1,2,(args,ctx)=>{const count=args[1]?integer(ctx.value(args[1])):1;if(count<0)fail('#VALUE!');return count?chars(text(ctx,args[0])).slice(-count).join(''):'';}),
 MID:fn(3,3,(args,ctx)=>{const start=integer(ctx.value(args[1])),count=integer(ctx.value(args[2]));if(start<1||count<0)fail('#VALUE!');return chars(text(ctx,args[0])).slice(start-1,start-1+count).join('');}),
 FIND:fn(2,3,(args,ctx)=>{const needle=text(ctx,args[0]),hay=text(ctx,args[1]),list=chars(hay),start=args[2]?integer(ctx.value(args[2])):1;if(start<1||start>list.length+1)fail('#VALUE!');ctx.work(hay.length);
  const at=hay.indexOf(needle,list.slice(0,start-1).join('').length);return at<0?fail('#VALUE!'):chars(hay.slice(0,at)).length+1;}),
 SEARCH:fn(2,3,(args,ctx)=>{const needle=text(ctx,args[0]),list=chars(text(ctx,args[1])),start=args[2]?integer(ctx.value(args[2])):1;if(start<1||start>list.length+1)fail('#VALUE!');
  if(!hasWildcard(needle)){ctx.work(list.length);const lower=list.join('').toLowerCase(),at=lower.indexOf(needle.toLowerCase(),list.slice(0,start-1).join('').length);return at<0?fail('#VALUE!'):chars(lower.slice(0,at)).length+1;}
  for(let i=start-1;i<=list.length;i++)if(wildcardMatch(list.slice(i).join(''),needle+'*',ctx))return i+1;return fail('#VALUE!');}),
 SUBSTITUTE:fn(3,4,(args,ctx)=>{const value=text(ctx,args[0]),from=text(ctx,args[1]),to=text(ctx,args[2]);if(!from)return value;
  if(!args[3])return limited(value.split(from).join(to),ctx);const nth=integer(ctx.value(args[3]));if(nth<1)fail('#VALUE!');let at=-1;for(let i=0;i<nth;i++){at=value.indexOf(from,at+1);if(at<0)return value;}return limited(value.slice(0,at)+to+value.slice(at+from.length),ctx);}),
 REPLACE:fn(4,4,(args,ctx)=>{const list=chars(text(ctx,args[0])),start=integer(ctx.value(args[1])),count=integer(ctx.value(args[2]));if(start<1||count<0)fail('#VALUE!');return limited(list.slice(0,start-1).join('')+text(ctx,args[3])+list.slice(start-1+count).join(''),ctx);}),
 TRIM:fn(1,1,(args,ctx)=>text(ctx,args[0]).replace(/ {2,}/g,' ').replace(/^ +| +$/g,'')),
 UPPER:fn(1,1,(args,ctx)=>text(ctx,args[0]).toUpperCase()),
 LOWER:fn(1,1,(args,ctx)=>text(ctx,args[0]).toLowerCase()),
 PROPER:fn(1,1,(args,ctx)=>{let previous=false;return chars(text(ctx,args[0])).map(char=>{const letter=/\p{L}/u.test(char),out=letter?(previous?char.toLowerCase():char.toUpperCase()):char;previous=letter;return out;}).join('');}),
 CONCAT:fn(1,MANY,(args,ctx)=>limited(args.flatMap(arg=>ctx.values(arg)).map(toText).join(''),ctx)),
 CONCATENATE:fn(1,MANY,(args,ctx)=>limited(args.flatMap(arg=>ctx.values(arg)).map(toText).join(''),ctx)),
 TEXTJOIN:fn(3,MANY,(args,ctx)=>{const delimiter=text(ctx,args[0]),skip=truth(ctx.value(args[1])),parts=args.slice(2).flatMap(arg=>ctx.values(arg)).map(toText).filter(value=>!skip||value!=='');return limited(parts.join(delimiter),ctx);}),
 TEXT:fn(2,2,(args,ctx)=>{const value=ctx.value(args[0]),code=text(ctx,args[1]);if(typeof value==='boolean')return toText(value);if(value===null)return formatSheetValue('0',code);
  const number=typeof value==='number'?value:value.trim()!==''&&Number.isFinite(Number(value))?Number(value):undefined;return number===undefined?value:formatSheetValue(numberText(number),code);}),
 VALUE:fn(1,1,(args,ctx)=>{const value=ctx.value(args[0]);if(typeof value==='number')return value;if(typeof value==='boolean')return fail('#VALUE!');const trimmed=(value??'').trim();
  if(/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(trimmed))return Number(trimmed.replaceAll(',',''));if(/%$/.test(trimmed)&&Number.isFinite(Number(trimmed.slice(0,-1)))&&trimmed.length>1)return Number(trimmed.slice(0,-1))/100;
  if(trimmed&&Number.isFinite(Number(trimmed)))return Number(trimmed);const date=parseSheetDateInput(trimmed);return date!==undefined?Number(date):fail('#VALUE!');}),
 EXACT:fn(2,2,(args,ctx)=>text(ctx,args[0])===text(ctx,args[1])),
 REPT:fn(2,2,(args,ctx)=>{const value=text(ctx,args[0]),count=integer(ctx.value(args[1]));if(count<0||value.length*count>TEXT_LIMIT)fail('#VALUE!');ctx.work(value.length*count);return value.repeat(count);}),
 // Date and time
 TODAY:fn(0,0,()=>nowSerial(false)),NOW:fn(0,0,()=>nowSerial(true)),
 DATE:fn(3,3,(args,ctx)=>dateSerial(num(ctx,args[0]),num(ctx,args[1]),num(ctx,args[2]))),
 DATEVALUE:fn(1,1,(args,ctx)=>{const value=ctx.value(args[0]);if(typeof value!=='string')return fail('#VALUE!');const date=parseSheetDateInput(value);return date!==undefined?Math.floor(Number(date)):fail('#VALUE!');}),
 YEAR:datePart('year'),MONTH:datePart('month'),DAY:datePart('day'),HOUR:datePart('hour'),MINUTE:datePart('minute'),SECOND:datePart('second'),
 WEEKDAY:fn(1,2,(args,ctx)=>{const serial=Math.floor(serialOf(ctx.value(args[0])));if(serial<0)fail('#NUM!');const type=args[1]?integer(ctx.value(args[1])):1,day=((serial-1)%7+7)%7;
  // Excel's serial 1 (1900-01-01) is a Sunday in its 1900 calendar; type 1 Sun=1…Sat=7, type 2 Mon=1…Sun=7, type 3 Mon=0…Sun=6.
  return type===1?day+1:type===2?(day+6)%7+1:type===3?(day+6)%7:fail('#NUM!');}),
 DAYS:fn(2,2,(args,ctx)=>Math.floor(serialOf(ctx.value(args[0])))-Math.floor(serialOf(ctx.value(args[1])))),
 DATEDIF:fn(3,3,(args,ctx)=>datedif(serialOf(ctx.value(args[0])),serialOf(ctx.value(args[1])),text(ctx,args[2]))),
 EDATE:fn(2,2,(args,ctx)=>addMonths(ctx.value(args[0]),num(ctx,args[1]),false)),
 EOMONTH:fn(2,2,(args,ctx)=>addMonths(ctx.value(args[0]),num(ctx,args[1]),true)),
};
/** Functions Excel stores with the `_xlfn.` prefix in OOXML; without it Excel shows #NAME? until re-entered. */
export const XLSX_FUTURE_FUNCTIONS=new Set(['IFS','SWITCH','XOR','CONCAT','TEXTJOIN','IFNA','DAYS','XLOOKUP','MAXIFS','MINIFS']);
/** Rewrites function names outside quoted text and sheet names; `formula` has no leading "=". */
function rewriteFunctionNames(formula:string,rename:(name:string)=>string){
 return formula.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'|([A-Za-z_][A-Za-z0-9_.]*)(?=\s*\()/g,(match,name:string|undefined)=>name===undefined?match:rename(name));
}
export const toXlsxFormula=(formula:string)=>rewriteFunctionNames(formula,name=>XLSX_FUTURE_FUNCTIONS.has(name.toUpperCase())?'_xlfn.'+name:name);
export const fromXlsxFormula=(formula:string)=>rewriteFunctionNames(formula,name=>name.replace(/^(?:_xlfn\.|_xlws\.)+/i,''));
