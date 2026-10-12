import {randomUUID} from 'node:crypto';
import * as Y from 'yjs';
import type {SQLiteProjectDatabase} from './database';

export interface QueueTarget {scope:string;server:string;workspace:string;doc:string;}
export class SQLiteSyncQueue {
 constructor(private database:SQLiteProjectDatabase){}
 private keys(t:QueueTarget){return [t.scope,t.server,t.workspace,t.doc];}
 list(t:QueueTarget){return this.database.db.prepare('SELECT seq,data,created,attempts,retryAt,error,generation FROM pending_updates WHERE scope=? AND server=? AND workspace=? AND doc=? ORDER BY seq').all(...this.keys(t)) as {seq:number;data:Uint8Array;created:number;attempts:number;retryAt:number;error:string|null;generation:string}[];}
 enqueue(t:QueueTarget,data:Uint8Array,snapshot?:Uint8Array){
  const check=new Y.Doc();try{Y.applyUpdate(check,data);if(snapshot)Y.applyUpdate(check,snapshot);}finally{check.destroy();}
  return this.database.db.transaction(()=>{
   const db=this.database.db,generation=this.database.generation;
   const previous=db.prepare('SELECT id,data FROM drafts WHERE scope=? AND server=? AND workspace=? AND doc=? AND generation=?').get(...this.keys(t),generation) as {id:string;data:Uint8Array}|undefined;
   const merged=Y.mergeUpdates([...(previous?[previous.data]:[]),...(snapshot?[snapshot]:[]),data]);
   // Persist a monotonic allocator: a delayed confirmation must not delete a reused row ID.
   const allocator=db.prepare("SELECT value FROM preferences WHERE scope='__system' AND key='sync-queue-sequence'").get() as {value:string}|undefined;
   const highest=db.prepare('SELECT coalesce(max(seq),0) n FROM pending_updates').get() as {n:number};
   const seq=Math.max(Number(allocator?.value??0),highest.n)+1;
   if(!Number.isSafeInteger(seq))throw new Error('同步队列序号超限');
   this.database.preference('__system','sync-queue-sequence',String(seq));
   db.prepare('INSERT INTO pending_updates(seq,scope,server,workspace,doc,data,created,generation) VALUES (?,?,?,?,?,?,?,?)').run(seq,...this.keys(t),Buffer.from(data),Date.now(),generation);
   if(previous)db.prepare('UPDATE drafts SET data=?,created=? WHERE id=?').run(Buffer.from(merged),Date.now(),previous.id);
   else db.prepare('INSERT INTO drafts(id,scope,server,workspace,doc,title,summary,data,created,generation) VALUES (?,?,?,?,?,?,?,?,?,?)').run(randomUUID(),...this.keys(t),'','',Buffer.from(merged),Date.now(),generation);
   return seq;
  }).immediate();
 }
 acknowledge(t:QueueTarget,seq:number){return this.database.db.transaction(()=>{
  const db=this.database.db;
  const row=db.prepare('SELECT 1 FROM pending_updates WHERE scope=? AND server=? AND workspace=? AND doc=? AND seq=? AND generation=?').get(...this.keys(t),seq,this.database.generation);
  if(!row)return;
  db.prepare('DELETE FROM pending_updates WHERE scope=? AND server=? AND workspace=? AND doc=? AND seq=? AND generation=?').run(...this.keys(t),seq,this.database.generation);
  this.database.preference(t.scope,'sync-last-success:'+JSON.stringify([t.server,t.workspace]),String(Date.now()));
  if(!this.list(t).some(row=>row.generation===this.database.generation))db.prepare('DELETE FROM drafts WHERE scope=? AND server=? AND workspace=? AND doc=? AND generation=?').run(...this.keys(t),this.database.generation);
 }).immediate();}
 fail(t:QueueTarget,seq:number,error:string){this.database.db.prepare('UPDATE pending_updates SET attempts=attempts+1,retryAt=?,error=? WHERE scope=? AND server=? AND workspace=? AND doc=? AND seq=? AND generation=?').run(Date.now()+Math.min(60000,1000*2**Math.min(this.list(t).find(r=>r.seq===seq)?.attempts??0,6)),error.slice(0,1000),...this.keys(t),seq,this.database.generation);}
}
