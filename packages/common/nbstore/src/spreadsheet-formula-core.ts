/** Shared formula primitives: the parser/calculator (spreadsheet-formula) and the function library both use these. */
export type ErrorCode='#VALUE!'|'#DIV/0!'|'#REF!'|'#NAME?'|'#NUM!'|'#N/A'|'#CYCLE!'|'#LIMIT!';
export const FORMULA_ERRORS:readonly ErrorCode[]=['#VALUE!','#DIV/0!','#REF!','#NAME?','#NUM!','#N/A','#CYCLE!','#LIMIT!'];
export class FormulaError extends Error {constructor(readonly code:ErrorCode){super(code);}}
export type Scalar=number|string|boolean|null;
export type Reference={sheet?:string;row:number;column:number;absoluteRow?:boolean;absoluteColumn?:boolean};
export type Ast={kind:'error';code:ErrorCode}|{kind:'literal';value:Scalar;omitted?:true}|{kind:'ref';ref:Reference}|{kind:'range';from:Reference;to:Reference}|{kind:'unary';op:string;value:Ast}|{kind:'binary';op:string;left:Ast;right:Ast}|{kind:'call';name:string;args:Ast[]};
export const fail=(code:ErrorCode):never=>{throw new FormulaError(code);};
export function finite(value:number){return Number.isFinite(value)?value:fail('#NUM!');}
export function numeric(value:Scalar){if(value===null||value==='')return 0;if(typeof value==='boolean')return value?1:0;const number=Number(value);if(!Number.isFinite(number))return fail('#VALUE!');return number;}
export function scalar(value:Scalar|Scalar[]):Scalar{if(Array.isArray(value))return fail('#VALUE!');return value;}
/** Lazily read 2-D view of a reference or range; `get` may throw the referenced cell's FormulaError. */
export interface FormulaMatrix {rows:number;columns:number;get(row:number,column:number):Scalar;}
export interface FormulaContext {
 /** Evaluate one argument to a single value (a range here is #VALUE!). */
 value(node:Ast):Scalar;
 /** Evaluate one argument, expanding ranges row by row. */
 values(node:Ast):Scalar[];
 /** Read an argument as a matrix; `rows`/`columns` resize from its top-left cell (SUMIF's sum_range). */
 matrix(node:Ast,rows?:number,columns?:number):FormulaMatrix;
 tick():void;
 /** Charge text/pattern work so crafted formulas cannot stall a collaborator's browser. */
 work(units:number):void;
}
export interface FormulaFunction {min:number;max:number;run(args:Ast[],ctx:FormulaContext):Scalar;}
