import {canTransition} from '@mydoc/agent-core';
import type {SQLiteProjectDatabase} from '../../nbstore/server/database';
import {validateEvent,TERMINAL_STATUSES,UNCONFIRMED_RESULT} from '@mydoc/agent-contracts';
const TERMINAL_SQL=TERMINAL_STATUSES.map(status=>`'${status}'`).join(',');
import type {AgentEvent,AgentRun,EventPort,ToolCall} from '@mydoc/agent-contracts';
function text(value:string,max=200){if(typeof value!=='string'||!value.length||value.length>max)throw new Error('Invalid Agent field');}
/** One trusted account scope. Never accept scope from a model or HTTP request. */
export class SQLiteAgentStore {
 constructor(private readonly database:SQLiteProjectDatabase,private readonly scope:string){text(scope);}
 private session(id:string){const row=this.database.db.prepare('SELECT * FROM agent_sessions WHERE scope=? AND id=?').get(this.scope,id);if(!row)throw new Error('Agent session unavailable');return row;}
 createSession(id:string,workspace:string,title:string,now=Date.now()){
  text(id);text(workspace);text(title,500);
  this.database.db.prepare('INSERT INTO agent_sessions VALUES (?,?,?,?,?,?)').run(this.scope,id,workspace,title,now,now);
 }
 listSessions(workspace:string){return this.database.db.prepare('SELECT id,title,created,updated FROM agent_sessions WHERE scope=? AND workspace=? ORDER BY updated DESC,id LIMIT 100').all(this.scope,workspace);}
 enqueue(session:string,id:string,key:string,input:string,now=Date.now()):AgentRun{
  text(id);text(key);text(input,100000);
  return this.database.db.transaction(()=>{
   this.session(session);
   const old=this.database.db.prepare('SELECT * FROM agent_runs WHERE scope=? AND session=? AND request_key=?').get(this.scope,session,key) as {id:string;input:string}|undefined;
   if(old){if(old.input!==input)throw new Error('Request key conflict');return this.load(old.id)!;}
   this.database.db.prepare('INSERT INTO agent_runs VALUES (?,?,?,?,?,?,?,?,?)').run(this.scope,id,session,key,input,'queued',0,this.database.generation,now);
   return this.load(id)!;
  }).immediate();
 }
 queued(session:string){this.session(session);return this.database.db.prepare("SELECT id,input,created FROM agent_runs WHERE scope=? AND session=? AND status='queued' ORDER BY created,id LIMIT 100").all(this.scope,session);}
 load(id:string):AgentRun|null{
  const row=this.database.db.prepare('SELECT * FROM agent_runs WHERE scope=? AND id=?').get(this.scope,id) as {id:string;session:string;request_key:string;status:AgentRun['status'];generation:number}|undefined;
  return row?{runId:row.id,sessionId:row.session,requestKey:row.request_key,status:row.status,generation:row.generation}:null;
 }
 claim(id:string):AgentRun{
  return this.database.db.transaction(()=>{
   const run=this.load(id);if(!run||run.status!=='queued')throw new Error('Run not queued');
   const active=this.database.db.prepare("SELECT 1 FROM agent_runs WHERE scope=? AND session=? AND status NOT IN ('queued','completed','cancelled','failed','budget_exhausted','interrupted')").get(this.scope,run.sessionId);if(active)throw new Error('Session already running');
   this.fence(run);this.database.db.prepare("UPDATE agent_runs SET status='preparing',generation=generation+1 WHERE scope=? AND id=?").run(this.scope,id);return this.load(id)!;
  }).immediate();
 }
 assertLease(lease:AgentRun){const run=this.load(lease.runId);if(!run||run.generation!==lease.generation||TERMINAL_STATUSES.includes(run.status as never))throw new Error('Stale Agent lease');this.fence(run);}
 cancelQueued(id:string){return this.database.db.transaction(()=>{const run=this.load(id);if(!run||run.status!=='queued')return false;this.fence(run);this.database.db.prepare("UPDATE agent_runs SET status='cancelled',generation=generation+1 WHERE scope=? AND id=?").run(this.scope,id);return true;}).immediate();}
 interrupt(lease:AgentRun){this.database.db.transaction(()=>{
  const run=this.load(lease.runId);if(!run||run.generation!==lease.generation||TERMINAL_STATUSES.includes(run.status as never))return;
  this.fence(run);
  const last=this.database.db.prepare('SELECT COALESCE(MAX(run_seq),0) n FROM agent_events WHERE scope=? AND run=?').get(this.scope,run.runId) as {n:number};
  let seq=last.n;
  const pending=this.database.db.prepare('SELECT id FROM agent_tool_calls WHERE scope=? AND run=? AND result IS NULL').all(this.scope,run.runId) as {id:string}[];
  for(const call of pending)this.append(run,{protocolVersion:1,runId:run.runId,sessionId:run.sessionId,seq:++seq,timestamp:Date.now(),payload:{type:'message',message:{role:'tool',callId:call.id,status:'error',text:UNCONFIRMED_RESULT+'; runtime interrupted, inspect domain receipt before retry'}}});
  this.append(run,{protocolVersion:1,runId:run.runId,sessionId:run.sessionId,seq:++seq,timestamp:Date.now(),payload:{type:'status',status:'interrupted'}});
  this.database.db.prepare('UPDATE agent_runs SET generation=generation+1 WHERE scope=? AND id=?').run(this.scope,run.runId);
 }).immediate();}
 private fence(run:AgentRun){
  const row=this.database.db.prepare('SELECT db_generation FROM agent_runs WHERE scope=? AND id=?').get(this.scope,run.runId) as {db_generation:string};
  if(this.database.restoring||row.db_generation!==this.database.generation)throw new Error('Database generation changed');
 }
 events(run:AgentRun):EventPort{return {append:async(event)=>this.append(run,event)};}
 private append(lease:AgentRun,event:AgentEvent){
  validateEvent(event);
  this.database.db.transaction(()=>{
   const current=this.load(lease.runId);
   if(!current||current.generation!==lease.generation||current.sessionId!==event.sessionId||current.runId!==event.runId)throw new Error('Stale Agent lease');
   this.fence(current);
   const payload=JSON.stringify(event);
   const old=this.database.db.prepare('SELECT payload FROM agent_events WHERE scope=? AND run=? AND run_seq=?').get(this.scope,event.runId,event.seq) as {payload:string}|undefined;
   if(old){if(old.payload!==payload)throw new Error('Event replay conflict');return;}
   if(TERMINAL_STATUSES.includes(current.status as never))throw new Error('Run already terminal');
   if(event.payload.type==='status'&&event.payload.status!==current.status&&!canTransition(current.status,event.payload.status))throw new Error('Invalid persisted transition');
   const last=this.database.db.prepare('SELECT COALESCE(MAX(run_seq),0) n FROM agent_events WHERE scope=? AND run=?').get(this.scope,event.runId) as {n:number};if(event.seq!==last.n+1)throw new Error('Agent event sequence gap');
   const next=this.database.db.prepare('SELECT COALESCE(MAX(cursor),0)+1 n FROM agent_events WHERE scope=? AND session=?').get(this.scope,event.sessionId) as {n:number};
   this.database.db.prepare('INSERT INTO agent_events VALUES (?,?,?,?,?,?)').run(this.scope,event.sessionId,next.n,event.runId,event.seq,payload);
   if(event.payload.type==='message'){
    const message=event.payload.message;
    this.database.db.prepare('INSERT INTO agent_messages VALUES (?,?,?,?,?)').run(this.scope,event.sessionId,next.n,event.runId,JSON.stringify(message));
    if(message.role==='assistant')for(const call of message.calls)this.database.db.prepare('INSERT INTO agent_tool_calls VALUES (?,?,?,?,NULL)').run(this.scope,event.runId,call.id,JSON.stringify(call));
    if(message.role==='tool'){
     if(!this.database.db.prepare('SELECT 1 FROM agent_tool_calls WHERE scope=? AND run=? AND id=?').get(this.scope,event.runId,message.callId))throw new Error('Orphan tool result');
     this.database.db.prepare('UPDATE agent_tool_calls SET result=? WHERE scope=? AND run=? AND id=?').run(JSON.stringify(message),this.scope,event.runId,message.callId);
    }
   }
   if(event.payload.type==='status')this.database.db.prepare('UPDATE agent_runs SET status=? WHERE scope=? AND id=?').run(event.payload.status,this.scope,event.runId);
   this.database.db.prepare('UPDATE agent_sessions SET updated=? WHERE scope=? AND id=?').run(event.timestamp,this.scope,event.sessionId);
  }).immediate();
 }
 replay(session:string,after=0){this.session(session);if(!Number.isSafeInteger(after)||after<0)throw new Error('Invalid cursor');return this.database.db.prepare('SELECT cursor,payload FROM agent_events WHERE scope=? AND session=? AND cursor>? ORDER BY cursor LIMIT 200').all(this.scope,session,after);}
 /** Only the authoritative host startup owner may recover its abandoned runs. Not every DB open. */
 recoverActive(){return this.database.db.transaction(()=>{
  const runs=this.database.db.prepare("SELECT id FROM agent_runs WHERE scope=? AND status NOT IN ('queued','completed','cancelled','failed','budget_exhausted','interrupted')").all(this.scope) as {id:string}[];
  for(const {id} of runs)this.interrupt(this.load(id)!);
  this.database.db.prepare('DELETE FROM agent_approvals WHERE scope=?').run(this.scope);
 }).immediate();}
 /** Started calls without a confirmed result; recovery inspects their domain receipts instead of replaying. */
 unconfirmed(run:string){
  const rows=this.database.db.prepare('SELECT id,payload,result FROM agent_tool_calls WHERE scope=? AND run=? ORDER BY rowid').all(this.scope,run) as {id:string;payload:string;result:string|null}[];
  return rows.filter(row=>{if(!row.result)return true;const result=JSON.parse(row.result) as {status:string;text:string};return result.status==='error'&&result.text.startsWith(UNCONFIRMED_RESULT);}).map(row=>JSON.parse(row.payload) as ToolCall);
 }
 /** Terminal runs keep messages; streamed text deltas are redundant with the saved assistant message. */
 compact(run:string){this.database.db.transaction(()=>{
  const current=this.load(run);if(!current||!TERMINAL_STATUSES.includes(current.status as never))return;
  this.database.db.prepare("DELETE FROM agent_events WHERE scope=? AND run=? AND json_extract(payload,'$.payload.type')='text'").run(this.scope,run);
 }).immediate();}
 /** Bounded history: keep the newest terminal runs; cursors stay monotonic because only older rows go. */
 prune(session:string,keep=50){this.database.db.transaction(()=>{
  this.session(session);
  const old=this.database.db.prepare(`SELECT id FROM agent_runs WHERE scope=? AND session=? AND status IN (${TERMINAL_SQL}) ORDER BY created DESC,id DESC LIMIT -1 OFFSET ?`).all(this.scope,session,keep) as {id:string}[];
  for(const {id} of old){
   for(const table of ['agent_tool_calls','agent_approvals'])this.database.db.prepare(`DELETE FROM ${table} WHERE scope=? AND run=?`).run(this.scope,id);
   for(const table of ['agent_events','agent_messages'])this.database.db.prepare(`DELETE FROM ${table} WHERE scope=? AND session=? AND run=?`).run(this.scope,session,id);
   this.database.db.prepare('DELETE FROM agent_runs WHERE scope=? AND id=?').run(this.scope,id);
  }
 }).immediate();}
 deleteSession(id:string){this.database.db.transaction(()=>{
  this.session(id);
  if(this.database.db.prepare(`SELECT 1 FROM agent_runs WHERE scope=? AND session=? AND status NOT IN ('queued',${TERMINAL_SQL})`).get(this.scope,id))throw new Error('Session still running');
  for(const table of ['agent_tool_calls','agent_approvals'])this.database.db.prepare(`DELETE FROM ${table} WHERE scope=? AND run IN (SELECT id FROM agent_runs WHERE scope=? AND session=?)`).run(this.scope,this.scope,id);
  for(const table of ['agent_events','agent_messages','agent_changesets','agent_runs'])this.database.db.prepare(`DELETE FROM ${table} WHERE scope=? AND session=?`).run(this.scope,id);
  this.database.db.prepare('DELETE FROM agent_sessions WHERE scope=? AND id=?').run(this.scope,id);
 }).immediate();}
}
