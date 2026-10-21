import type * as Y from 'yjs';
import type {SpreadsheetCellDraft} from './workspace';
export const spreadsheetDraftPrefix=(id:string)=>'sheet-cell-draft:'+id+':';
/** Account-local Y.Map is persisted by the existing SQLite personal preferences storage.
 * Keep one entry per cell so independent drafts do not replace one another. */
export class SpreadsheetDraftStore extends Map<string,SpreadsheetCellDraft>{
 constructor(private preferences:Y.Map<unknown>,private resource:string){super();}
 private get prefix(){return spreadsheetDraftPrefix(this.resource);}
 override get size(){return [...this.keys()].length;}
 override has(key:string){return this.preferences.has(this.prefix+key);}
 override get(key:string){return this.preferences.get(this.prefix+key) as SpreadsheetCellDraft|undefined;}
 override set(key:string,value:SpreadsheetCellDraft){this.preferences.set(this.prefix+key,JSON.parse(JSON.stringify(value)));return this;}
 override delete(key:string){const exists=this.has(key);this.preferences.delete(this.prefix+key);return exists;}
 override clear(){for(const key of this.keys())this.delete(key);}
 override *keys():MapIterator<string>{for(const key of this.preferences.keys())if(key.startsWith(this.prefix))yield key.slice(this.prefix.length);}
 override *entries():MapIterator<[string,SpreadsheetCellDraft]>{for(const key of this.keys())yield [key,this.get(key)!];}
 override *values():MapIterator<SpreadsheetCellDraft>{for(const [,value]of this.entries())yield value;}
 override [Symbol.iterator](){return this.entries();}
 override forEach(callback:(value:SpreadsheetCellDraft,key:string,map:Map<string,SpreadsheetCellDraft>)=>void,thisArg?:unknown){for(const [key,value]of this.entries())callback.call(thisArg,value,key,this);}
}
