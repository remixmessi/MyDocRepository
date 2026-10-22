import type {SpreadsheetCellType} from './spreadsheet-cell-type';
import {FormulaError,FORMULA_ERRORS,fail,finite,numeric,scalar,type Ast,type ErrorCode,type FormulaContext,type FormulaMatrix,type Reference,type Scalar} from './spreadsheet-formula-core';
import {FORMULA_FUNCTIONS,TEXT_LIMIT,toText} from './spreadsheet-functions';
export type FormulaCellTypes=(sheet:string,row:number,column:number)=>SpreadsheetCellType;
export type FormulaGrid=ReadonlyArray<ReadonlyArray<string>>;
export type FormulaWorkbook=Record<string,FormulaGrid>;
/** `work` bounds text and wildcard processing across one calculation (pattern steps and produced characters). */
export const FORMULA_LIMITS={source:10000,tokens:2048,depth:128,range:100000,steps:200000,work:5_000_000};
type Token={kind:'error'|'number'|'string'|'name'|'sheet'|'symbol'|'end';text:string};
const errors=new Set<string>(FORMULA_ERRORS);
export const isFormulaError=(value:string)=>errors.has(value);

function tokenize(source:string):Token[]{
 if(source.length>FORMULA_LIMITS.source)fail('#LIMIT!');
 const out:Token[]=[];let i=0;
 while(i<source.length){
  if(/\s/.test(source[i])){i++;continue;}
  if(out.length>=FORMULA_LIMITS.tokens)fail('#LIMIT!');
  const quote=source[i];
  if(quote==='"'||quote==="'"){
   let text='',closed=false;i++;
   while(i<source.length){if(source[i]===quote){if(source[i+1]===quote){text+=quote;i+=2;continue;}i++;closed=true;break;}text+=source[i++];}
   if(!closed)fail('#VALUE!');out.push({kind:quote==='"'?'string':'sheet',text});continue;
  }
  const error=/^#(?:VALUE!|DIV\/0!|REF!|NAME\?|NUM!|N\/A|CYCLE!|LIMIT!)/i.exec(source.slice(i));if(error){out.push({kind:'error',text:error[0].toUpperCase()});i+=error[0].length;continue;}
  const number=/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(source.slice(i));
  if(number){out.push({kind:'number',text:number[0]});i+=number[0].length;continue;}
  const op=/^(?:<=|>=|<>|[+\-*/^&%=<>(),!:])/.exec(source.slice(i));
  if(op){out.push({kind:'symbol',text:op[0]});i+=op[0].length;continue;}
  const name=/^[\p{L}_$][\p{L}\p{N}_.$]*/u.exec(source.slice(i));
  if(name){out.push({kind:'name',text:name[0]});i+=name[0].length;continue;}
  fail('#VALUE!');
 }
 out.push({kind:'end',text:''});return out;
}
// Excel order: comparison < & < +- < */ < ^ < unary minus < %. Unary minus binds tighter than ^, so -2^2 is 4.
const precedence:Record<string,number>={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
const UNARY=6;
class Parser {
 private index=0;private depth=0;
 constructor(private tokens:Token[]){}
 private peek(){return this.tokens[this.index];}
 private take(){return this.tokens[this.index++];}
 private accept(value:string){if(this.peek().kind==='symbol'&&this.peek().text===value){this.index++;return true;}return false;}
 private require(value:string){if(!this.accept(value))fail('#VALUE!');}
 parse(){const node=this.expression();if(this.peek().kind!=='end')fail('#VALUE!');return node;}
 private expression(min=1):Ast{
  if(++this.depth>FORMULA_LIMITS.depth)fail('#LIMIT!');
  try{
   let node=this.primary();
   while(this.peek().kind==='symbol'&&(precedence[this.peek().text]??0)>=min){const op=this.take().text;node={kind:'binary',op,left:node,right:this.expression(precedence[op]+1)};}
   return node;
  }finally{this.depth--;}
 }
 private reference(first:Token):Reference{
  let sheet:string|undefined,cell=first;
  if(this.accept('!')){sheet=first.text;cell=this.take();}
  const match=/^(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)$/i.exec(cell.text);
  if(!match||cell.kind!=='name')return fail('#REF!');
  const row=Number(match[4])-1,column=columnNumber(match[2]);
  if(row>=1048576||column>=16384)fail('#REF!');
  return {sheet,row,column,absoluteRow:!!match[3],absoluteColumn:!!match[1]};
 }
 private primary():Ast{
  if(this.accept('+'))return {kind:'unary',op:'+',value:this.expression(UNARY)};
  if(this.accept('-'))return {kind:'unary',op:'-',value:this.expression(UNARY)};
  let node=this.atom();
  // Postfix percent binds tightest: 50% is 0.5 and -50% is -0.5.
  while(this.accept('%'))node={kind:'unary',op:'%',value:node};
  return node;
 }
 private atom():Ast{
  if(this.accept('(')){const node=this.expression();this.require(')');return node;}
  const token=this.take();
  if(token.kind==='error')return {kind:'error',code:token.text as ErrorCode};
  if(token.kind==='string')return {kind:'literal',value:token.text};
  if(token.kind==='name'&&this.accept('(')){
   const args:Ast[]=[];
   // An omitted argument (VLOOKUP(x,r,2,) or XLOOKUP(…,,0,-1)) evaluates as blank; functions can tell it apart.
   const omitted=()=>this.peek().kind==='symbol'&&[',',')'].includes(this.peek().text);
   if(!this.accept(')')){do{args.push(omitted()?{kind:'literal',value:null,omitted:true}:this.expression());}while(this.accept(','));this.require(')');}
   // Excel files prefix newer functions (_xlfn.XLOOKUP); the stored formula keeps whatever the user typed.
   return {kind:'call',name:token.text.toUpperCase().replace(/^(?:_XLFN\.|_XLWS\.)+/,''),args};
  }
  if(token.kind==='name'&&['TRUE','FALSE'].includes(token.text.toUpperCase())&&this.peek().text!=='!')return {kind:'literal',value:token.text.toUpperCase()==='TRUE'};
  if(token.kind==='number'&&this.peek().text!=='!')return {kind:'literal',value:finite(Number(token.text))};
  if(!['name','sheet','number'].includes(token.kind))return fail('#VALUE!');
  if(token.kind==='name'&&this.peek().text!=='!'&&!/^\$?[A-Z]{1,3}\$?[1-9]\d*$/i.test(token.text))return fail('#NAME?');
  const from=this.reference(token);
  if(!this.accept(':'))return {kind:'ref',ref:from};
  const to=this.reference(this.take());
  if(to.sheet&&from.sheet&&to.sheet.toLowerCase()!==from.sheet.toLowerCase())fail('#REF!');
  if(to.sheet&&!from.sheet)fail('#REF!');
  return {kind:'range',from,to:{...to,sheet:to.sheet??from.sheet}};
 }
}
function parse(source:string){return new Parser(tokenize(source)).parse();}
function columnNumber(value:string){let n=0;for(const char of value.toUpperCase())n=n*26+char.charCodeAt(0)-64;return n-1;}
function cellKey(sheet:string,row:number,column:number){return JSON.stringify([sheet,row,column]);}
function resolveSheet(book:FormulaWorkbook,name:string){const key=Object.keys(book).find(key=>key.toLowerCase()===name.toLowerCase());if(key===undefined)return fail('#REF!');return key;}
function* rangeReferences(node:Extract<Ast,{kind:'range'}>){
 const r0=Math.min(node.from.row,node.to.row),r1=Math.max(node.from.row,node.to.row),c0=Math.min(node.from.column,node.to.column),c1=Math.max(node.from.column,node.to.column);
 if((r1-r0+1)*(c1-c0+1)>FORMULA_LIMITS.range)fail('#LIMIT!');
 for(let row=r0;row<=r1;row++)for(let column=c0;column<=c1;column++)yield {sheet:node.from.sheet,row,column};
}
class Calculation {
 private steps=0;private effort=0;private cache=new Map<string,Scalar>();private active=new Set<string>();private depth=0;
 constructor(readonly book:FormulaWorkbook,readonly types?:FormulaCellTypes){}
 private tick(){if(++this.steps>FORMULA_LIMITS.steps)fail('#LIMIT!');}
 private work(units:number){this.effort+=units;if(this.effort>FORMULA_LIMITS.work)fail('#LIMIT!');}
 /** A reference or range as a lazily evaluated matrix; any other expression is a 1×1 matrix of its value. */
 private matrix(node:Ast,sheet:string,rows?:number,columns?:number):FormulaMatrix{
  if(node.kind!=='ref'&&node.kind!=='range'){const value=scalar(this.evaluate(node,sheet));if((rows??1)!==1||(columns??1)!==1)fail('#VALUE!');return {rows:1,columns:1,get:()=>value};}
  const from=node.kind==='ref'?node.ref:node.from,to=node.kind==='ref'?node.ref:node.to,top=Math.min(from.row,to.row),left=Math.min(from.column,to.column);
  const height=rows??Math.max(from.row,to.row)-top+1,width=columns??Math.max(from.column,to.column)-left+1;
  if(height*width>FORMULA_LIMITS.range)fail('#LIMIT!');if(top+height>1048576||left+width>16384)fail('#REF!');
  return {rows:height,columns:width,get:(r,c)=>this.cell({sheet:from.sheet,row:top+r,column:left+c},sheet)};
 }
 private context(sheet:string):FormulaContext{
  return {value:node=>scalar(this.evaluate(node,sheet)),values:node=>{const value=this.evaluate(node,sheet);return Array.isArray(value)?value:[value];},matrix:(node,rows,columns)=>this.matrix(node,sheet,rows,columns),tick:()=>this.tick(),work:units=>this.work(units)};
 }
 cell(ref:Reference,current:string):Scalar{
  this.tick();const sheet=ref.sheet?resolveSheet(this.book,ref.sheet):current,key=cellKey(sheet,ref.row,ref.column);
  if(this.active.has(key))return fail('#CYCLE!');
  if(this.cache.has(key))return this.cache.get(key)!;
  const raw=this.book[sheet]?.[ref.row]?.[ref.column]??'';
  let value:Scalar=null;
  if(this.types?.(sheet,ref.row,ref.column)==='text'){value=raw;}else if(raw.trim().startsWith('=')){
   if(this.active.size>=FORMULA_LIMITS.depth)fail('#LIMIT!');
   this.active.add(key);try{value=scalar(this.evaluate(parse(raw.trim().slice(1)),sheet));}finally{this.active.delete(key);}
  }else if(raw!==''){const number=Number(raw);value=raw.trim()!==''&&Number.isFinite(number)?number:raw;}
  this.cache.set(key,value);return value;
 }
 evaluate(node:Ast,sheet:string):Scalar|Scalar[]{
  this.tick();if(++this.depth>FORMULA_LIMITS.depth)fail('#LIMIT!');
  try{
   switch(node.kind){
    case 'error':return fail(node.code);
    case 'literal':return node.value;
    case 'ref':return this.cell(node.ref,sheet);
    case 'range':return [...rangeReferences(node)].map(ref=>this.cell(ref,sheet));
    case 'unary':{const value=numeric(scalar(this.evaluate(node.value,sheet)));return finite(node.op==='-'?-value:node.op==='%'?value/100:value);}
    case 'binary':{
     const a=scalar(this.evaluate(node.left,sheet)),b=scalar(this.evaluate(node.right,sheet));
     if(node.op==='&'){const text=toText(a)+toText(b);this.work(text.length);return text.length>TEXT_LIMIT?fail('#VALUE!'):text;}
     if(['=','<>','<','>','<=','>='].includes(node.op)){
      const left=typeof a==='string'?a.toLowerCase():a??0,right=typeof b==='string'?b.toLowerCase():b??0;
      const comparison=typeof left===typeof right?(left===right?0:left<right?-1:1):(String(left)<String(right)?-1:1);
      switch(node.op){case '=':return comparison===0;case '<>':return comparison!==0;case '<':return comparison<0;case '>':return comparison>0;case '<=':return comparison<=0;default:return comparison>=0;}
     }
     const left=numeric(a),right=numeric(b);
     switch(node.op){case '+':return finite(left+right);case '-':return finite(left-right);case '*':return finite(left*right);case '^':if(left===0&&right<0)return fail('#DIV/0!');return finite(left**right);default:if(right===0)return fail('#DIV/0!');return finite(left/right);}
    }
    case 'call':return this.call(node,sheet);
   }
  }finally{this.depth--;}
 }
 private call(node:Extract<Ast,{kind:'call'}>,sheet:string):Scalar{
  const entry=FORMULA_FUNCTIONS[node.name];if(!entry)return fail('#NAME?');
  if(node.args.length<entry.min||node.args.length>entry.max)fail('#VALUE!');
  return entry.run(node.args,this.context(sheet));
 }
}
/** Workbook grids are immutable snapshots. Values are derived; shared input is never overwritten. */
export function evaluateFormula(input:string,grid:FormulaGrid,workbook?:FormulaWorkbook,types?:FormulaCellTypes):string{
 if(!input.trim().startsWith('='))return input;
 try{
  const current=Object.entries(workbook??{}).find(([,value])=>value===grid)?.[0]??'__current__';
  const book:FormulaWorkbook=Object.assign(Object.create(null),workbook??{},{[current]:grid});
  const result=scalar(new Calculation(book,types).evaluate(parse(input.trim().slice(1)),current));
  return typeof result==='boolean'?result?'TRUE':'FALSE':String(result??0);
 }catch(error){return error instanceof FormulaError?error.code:'#VALUE!';}
}
function dependencies(node:Ast,limit:number):Reference[]{
 const refs:Reference[]=[];const add=(ref:Reference)=>{if(refs.length>=limit)fail('#LIMIT!');refs.push(ref);};
 const visit=(value:Ast):void=>{switch(value.kind){case 'ref':add(value.ref);break;case 'range':for(const ref of rangeReferences(value))add(ref);break;case 'unary':visit(value.value);break;case 'binary':visit(value.left);visit(value.right);break;case 'call':for(const arg of value.args)visit(arg);break;}};visit(node);return refs;
}
/** Conservative dependency graph: both IF branches are checked, string literals are not references. */
export function detectFormulaCycles(workbook:FormulaWorkbook,types?:FormulaCellTypes,selectedSheets?:string[]){
 const visiting=new Set<string>(),done=new Set<string>(),cycles=new Set<string>();
 let steps=0;
 type Frame={sheet:string;row:number;column:number;exit?:boolean};
 for(const [sheet,grid]of Object.entries(workbook).filter(([sheet])=>!selectedSheets||selectedSheets.includes(sheet)))for(let row=0;row<grid.length;row++)for(let column=0;column<grid[row].length;column++){
  const stack:Frame[]=[{sheet,row,column}];
  while(stack.length){
   if(++steps>FORMULA_LIMITS.steps)fail('#LIMIT!');
   const frame=stack.pop()!,key=cellKey(frame.sheet,frame.row,frame.column);
   if(frame.exit){visiting.delete(key);done.add(key);continue;}
   if(visiting.has(key)){cycles.add(`${frame.sheet}!${frame.row}:${frame.column}`);continue;}
   if(done.has(key))continue;
   const value=workbook[frame.sheet]?.[frame.row]?.[frame.column]??'';
   if(types?.(frame.sheet,frame.row,frame.column)==='text'||!value.trim().startsWith('=')){done.add(key);continue;}
   let refs:Reference[];try{refs=dependencies(parse(value.trim().slice(1)),FORMULA_LIMITS.steps-steps);}catch(error){if(error instanceof FormulaError&&error.code==='#LIMIT!')throw error;done.add(key);continue;}
   visiting.add(key);stack.push({...frame,exit:true});
   for(const ref of refs.reverse()){
    let target=frame.sheet;try{if(ref.sheet)target=resolveSheet(workbook,ref.sheet);}catch{continue;}
    stack.push({sheet:target,row:ref.row,column:ref.column});
   }
  }
 }
 return [...cycles];
}

// 3: function library (lookup, conditional aggregation, text, date), & ^ % operators and #N/A.
export const FORMULA_CALCULATION_VERSION='3';
/** One bounded calculation context and cache for the whole immutable grid. */
export function calculateFormulaWorkbookTyped(book:FormulaWorkbook,types?:FormulaCellTypes,selectedSheets=Object.keys(book)):Record<string,Array<Array<number|string|boolean|null>>>{
 const address=(row:number,column:number)=>{let letters='';for(let n=column+1;n;n=Math.floor((n-1)/26))letters=String.fromCharCode(65+(n-1)%26)+letters;return letters+(row+1);};
 let cycles:string[];try{cycles=detectFormulaCycles(book,types,selectedSheets);}catch{throw new Error('公式依赖检查超过计算上限，请简化表格或选择公式原文导出');}
 if(cycles.length)throw new Error('检测到公式循环依赖：'+cycles.slice(0,10).map(value=>{const match=/^(.*)!(\d+):(\d+)$/.exec(value);return match?(Object.keys(book).length>1?match[1]+'!':'')+address(Number(match[2]),Number(match[3])):value;}).join('、')+'。无法导出计算结果，请修正公式或选择公式原文');
 const calculation=new Calculation(book,types),errors:string[]=[];
 const values:Record<string,Array<Array<number|string|boolean|null>>>=Object.create(null);
 for(const sheet of selectedSheets){const grid=book[sheet];if(!grid)throw new Error('工作表不存在');values[sheet]=grid.map((row,r)=>row.map((raw,c)=>{
  if(types?.(sheet,r,c)==='text'||!raw.trim().startsWith('='))return raw;
  try{return calculation.cell({row:r,column:c},sheet);}
  catch(error){const code=error instanceof FormulaError?error.code:'#VALUE!';errors.push((Object.keys(book).length>1?sheet+'!':'')+address(r,c)+' '+code);if(code==='#LIMIT!'||errors.length>=20)throw new Error('无法导出计算结果：'+errors.join('；')+'。请修正公式或选择公式原文');return '';}
 }));}
 if(errors.length)throw new Error('无法导出计算结果：'+errors.join('；')+'。请修正公式或选择公式原文');return values;
}

export function calculateFormulaGridTyped(grid:FormulaGrid,types?:FormulaCellTypes,sheet='当前工作表'):Array<Array<number|string|boolean|null>>{return calculateFormulaWorkbookTyped({[sheet]:grid},types)[sheet];}

/** Display/export formatting must not be used to infer formula result types. */
export function calculateFormulaGrid(grid:FormulaGrid,types?:FormulaCellTypes,sheet='当前工作表'):string[][]{
 return calculateFormulaGridTyped(grid,types,sheet).map(row=>row.map(value=>typeof value==='boolean'?value?'TRUE':'FALSE':String(value??0)));
}

/** Unsupported syntax is conservatively treated as a possible reference before destructive changes. */
export function formulaMayReferenceSheet(input:string,sheet:string){
 if(!input.trim().startsWith('='))return false;
 const same=(name:string|undefined)=>name?.toLowerCase()===sheet.toLowerCase();
 const visit=(node:Ast):boolean=>{switch(node.kind){case 'ref':return same(node.ref.sheet);case 'range':return same(node.from.sheet)||same(node.to.sheet);case 'binary':return visit(node.left)||visit(node.right);case 'unary':return visit(node.value);case 'call':return node.args.some(visit);default:return false;}};
 try{return visit(parse(input.trim().slice(1)));}catch{return true;}
}

/** One bounded cache for all visible results, including cell-local error codes. */
export function previewFormulaWorkbook(book:FormulaWorkbook,types?:FormulaCellTypes):Record<string,string[][]>{
 const calculation=new Calculation(book,types),result:Record<string,string[][]>=Object.create(null);
 for(const [sheet,grid]of Object.entries(book))result[sheet]=grid.map((row,r)=>row.map((raw,c)=>{
  if(!raw.trim().startsWith('=')||types?.(sheet,r,c)==='text')return raw;
  try{const value=calculation.cell({row:r,column:c},sheet);return typeof value==='boolean'?value?'TRUE':'FALSE':String(value??0);}catch(error){return error instanceof FormulaError?error.code:'#VALUE!';}
 }));return result;
}

export type FormulaStructureChange={sheet:string;axis:'row'|'column';index:number;delete:boolean};
/** Parse, transform reference nodes, and serialize; quoted strings are never rewritten as addresses. */
export function transformFormulaReferences(input:string,currentSheet:string,change:FormulaStructureChange){
 if(!input.trim().startsWith('='))return input;let changed=false;
 const same=(ref:Reference)=>(ref.sheet??currentSheet).toLowerCase()===change.sheet.toLowerCase(),axis=change.axis;
 const single=(ref:Reference):Reference|undefined=>{if(!same(ref))return ref;const n=ref[axis];if(change.delete&&n===change.index){changed=true;return undefined;}const next=n+(change.delete?(n>change.index?-1:0):(n>=change.index?1:0));if(next>= (axis==='row'?1048576:16384))throw new Error('公式引用变换超出工作表边界');if(next!==n)changed=true;return {...ref,[axis]:next};};
 const visit=(node:Ast):Ast=>{switch(node.kind){
  case 'ref':{const ref=single(node.ref);return ref?{kind:'ref',ref}:{kind:'error',code:'#REF!'};}
  case 'range':{
   if(!same(node.from))return node;const a=node.from[axis],b=node.to[axis],low=Math.min(a,b),high=Math.max(a,b);
   if(change.delete&&change.index>=low&&change.index<=high){changed=true;if(low===high)return {kind:'error',code:'#REF!'};const nextLow=low,nextHigh=high-1;return {kind:'range',from:{...node.from,[axis]:a<=b?nextLow:nextHigh},to:{...node.to,[axis]:a<=b?nextHigh:nextLow}};}
   const from=single(node.from),to=single(node.to);return from&&to?{kind:'range',from,to}:{kind:'error',code:'#REF!'};
  }
  case 'unary':return {...node,value:visit(node.value)};
  case 'binary':return {...node,left:visit(node.left),right:visit(node.right)};
  case 'call':return {...node,args:node.args.map(visit)};
  default:return node;
 }};
 const address=(ref:Reference,qualify=true)=>{let column='';for(let n=ref.column+1;n;n=Math.floor((n-1)/26))column=String.fromCharCode(65+(n-1)%26)+column;return (qualify&&ref.sheet?"'"+ref.sheet.replaceAll("'","''")+"'!":'')+(ref.absoluteColumn?'$':'')+column+(ref.absoluteRow?'$':'')+(ref.row+1);};
 const serialize=(node:Ast):string=>{switch(node.kind){case 'error':return node.code;case 'literal':return node.omitted?'':typeof node.value==='string'?'"'+node.value.replaceAll('"','""')+'"':typeof node.value==='boolean'?node.value?'TRUE':'FALSE':String(node.value??0);case 'ref':return address(node.ref);case 'range':return address(node.from)+':'+address(node.to,false);case 'unary':return node.op==='%'?'('+serialize(node.value)+')%':node.op+'('+serialize(node.value)+')';case 'binary':return '('+serialize(node.left)+node.op+serialize(node.right)+')';case 'call':return node.name+'('+node.args.map(serialize).join(',')+')';}};
 // Functions outside the library (INDIRECT, OFFSET, ROW…) may build references from text or position, which a
 // rewrite cannot follow, so they block structure changes just like unparsable formulas.
 const known=(node:Ast):boolean=>{switch(node.kind){case 'call':return !!FORMULA_FUNCTIONS[node.name]&&node.args.every(known);case 'unary':return known(node.value);case 'binary':return known(node.left)&&known(node.right);default:return true;}};
 let ast:Ast;try{const parsed=parse(input.trim().slice(1));if(!known(parsed))throw new Error('unknown');ast=visit(parsed);}catch{throw new Error('公式不支持安全引用变换，请先修正公式或改为文本类型');}if(!changed)return input;const result='='+serialize(ast);if(result.length>FORMULA_LIMITS.source)throw new Error('变换后的公式超过长度限制');return result;
}
