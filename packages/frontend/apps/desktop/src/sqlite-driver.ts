import {DatabaseSync,type SQLInputValue} from 'node:sqlite';
import type {SQLiteConnection} from '../../../../common/nbstore/server/database';
// Electron ships SQLite with Node: no platform-specific addon rebuild is needed.
export function connectDesktopDatabase(path:string):SQLiteConnection {
 const db=new DatabaseSync(path);let depth=0;
 return {
  exec:sql=>db.exec(sql),
  pragma(sql,options){const rows=db.prepare('PRAGMA '+sql).all();return options?.simple?Object.values(rows[0]??{})[0]:rows;},
  prepare(sql){const stmt=db.prepare(sql);return {get:(...args)=>stmt.get(...args as SQLInputValue[]),all:(...args)=>stmt.all(...args as SQLInputValue[]),run:(...args)=>stmt.run(...args as SQLInputValue[])};},
  transaction<T>(fn:()=>T){return {immediate(){const nested=depth>0,name='mydoc_'+depth;db.exec(nested?'SAVEPOINT '+name:'BEGIN IMMEDIATE');depth++;try{const result=fn();db.exec(nested?'RELEASE '+name:'COMMIT');return result;}catch(error){db.exec(nested?'ROLLBACK TO '+name:'ROLLBACK');if(nested)db.exec('RELEASE '+name);throw error;}finally{depth--;}}};},
  close:()=>db.close(),
 };
}
