import {sheetColumnIdentity} from './sheet-columns';
import type {SpreadsheetCellType} from './spreadsheet-cell-type';
import * as Y from 'yjs';
export interface SheetCellEdit {row:number;column:number;value:string;cellType?:SpreadsheetCellType;base:string;}
export const MAX_SHEET_BATCH_CELLS=500;
/** `format` on a target is merged over the operation's shared patch (border presets differ per cell). */
/** One row/column of a `setDimensions` operation; null clears the field (back to the default size, or shown). */
export interface SheetDimensionTarget {key:string;size?:number|null;hidden?:true|null;}
export interface SheetFormatTarget {rowKey:string;columnKey:string;format?:SheetFormatPatch;}
export type SheetFormatPatch=Partial<Record<'bold'|'italic'|'underline'|'strike'|'wrap',true|null>&Record<'color'|'fill'|'align'|'verticalAlign'|'numberFormat'|'borderTop'|'borderRight'|'borderBottom'|'borderLeft',string|null>&Record<'fontSize',number|null>>;
export interface SheetOperation {operationId:string;officeGeneration?:string;rowKey?:string;columnKey?:string;sheet?:string;type:'insertRow'|'deleteRow'|'insertColumn'|'deleteColumn'|'setCell'|'appendRow'|'deleteLastRow'|'setCells'|'addSheet'|'deleteSheet'|'moveSheetLeft'|'moveSheetRight'|'renameSheet'|'setFormat'|'mergeCells'|'unmergeCells'|'putSheetObject'|'deleteSheetObject'|'setDimensions';objectId?:string;axis?:'row'|'column';dimensions?:SheetDimensionTarget[];object?:import('./spreadsheet-objects').SheetObject;name?:string;cells?:SheetCellEdit[];targets?:SheetFormatTarget[];format?:SheetFormatPatch;rowSpan?:number;colSpan?:number;row?:number;column?:number;value?:string;cellType?:SpreadsheetCellType;base:string;}
/** Includes Yjs item identity, so changing a cell away and back still invalidates its old revision. */
export function cellRevision(sheet:Y.Array<Y.Array<string>>,row:number,column:number){
 const target=sheet.get(row);if(!target)throw new Error('行已不存在');
 return JSON.stringify([Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(target,column))),target.get(column)??null,sheetColumnIdentity(sheet,column)]);
}
export function structureRevision(doc:Y.Doc){return btoa(Array.from(Y.encodeStateVector(doc),value=>String.fromCharCode(value)).join(''));}

export function cellDraftKey(sheet:Y.Array<Y.Array<string>>,row:number,column:number){return JSON.stringify([Array.from(Y.encodeRelativePosition(Y.createRelativePositionFromTypeIndex(sheet,row))),sheetColumnIdentity(sheet,column)]);}
