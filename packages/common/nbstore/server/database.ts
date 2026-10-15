import {randomUUID} from 'node:crypto';
import {mkdirSync} from 'node:fs';
import {dirname} from 'node:path';
import * as Y from 'yjs';
import {SCHEMA_VERSION,SCHEMA_MIGRATIONS,SCHEMA_COLUMNS,SCHEMA_CREATE_SQL,BUSINESS_TABLES} from '../src/schema';
export interface SQLiteConnection {
 exec(sql:string):unknown;
 pragma(sql:string,options?:{simple?:boolean}):unknown;
 prepare(sql:string):{get(...values:unknown[]):unknown;all(...values:unknown[]):unknown[];run(...values:unknown[]):unknown};
 transaction<T>(fn:()=>T):{immediate:()=>T};
 close():void;
}
export interface StartupProtection {
 /** Runs before any migration or schema writes; a failed snapshot stops startup. */
 beforeMigration?:(db:SQLiteConnection)=>void;
 preserve?:Record<string,string>;
}
export class SQLiteProjectDatabase {
 restoring=false;
 get generation(){let row=this.db.prepare("SELECT value FROM preferences WHERE scope='__system' AND key='restore-generation'").get() as {value:string}|undefined;if(!row){this.db.prepare("INSERT OR IGNORE INTO preferences VALUES ('__system','restore-generation',?)").run(randomUUID());row=this.db.prepare("SELECT value FROM preferences WHERE scope='__system' AND key='restore-generation'").get() as {value:string};}return row.value;}
 set generation(value:string){this.db.prepare("INSERT INTO preferences VALUES ('__system','restore-generation',?) ON CONFLICT(scope,key) DO UPDATE SET value=excluded.value").run(value);}
 readonly db:SQLiteConnection;
 readonly initializedThisStart:boolean;
 readonly schemaVersion=SCHEMA_VERSION;
 constructor(readonly path:string,readonly connect:(path:string)=>SQLiteConnection,upgrade?:(db:SQLiteConnection)=>void,protection:StartupProtection={}){
  mkdirSync(dirname(path),{recursive:true});this.db=connect(path);
  try {
   this.db.pragma('busy_timeout = 5000');
   const initialVersion=this.db.pragma('user_version',{simple:true}) as number;
   if(initialVersion>this.schemaVersion)throw new Error('数据库版本高于当前程序，请升级程序后再启动');
   protection.beforeMigration?.(this.db);
   // Recheck under the write lock: concurrent starts must not initialize twice.
   // 逐版本迁移：已有库只补做缺少的步骤，整体在一个写事务内，失败回滚且 user_version 不前进。
   this.initializedThisStart=this.db.transaction(()=>{
    const version=this.db.pragma('user_version',{simple:true}) as number;
    if(version>this.schemaVersion)throw new Error('数据库版本高于当前程序，请升级程序后再启动');
    if(!Number.isSafeInteger(version)||version<0)throw new Error('数据库版本无效');
    // Keep exact original values inside the same write transaction, including
    // binary Yjs updates and attachments. A destructive migration must roll back.
    const retained=retainRows(this.db,{...Object.fromEntries(BUSINESS_TABLES),...protection.preserve});
    for(let next=version+1;next<=this.schemaVersion;next++){
     const migration=SCHEMA_MIGRATIONS[next];
     if(!migration)throw new Error('缺少数据库结构迁移步骤 '+next);
     this.db.exec(migration);
    }
    // 即使版本已经最新，也检查并补建缺失表和索引。补建语句必须保持幂等。
    this.db.exec(SCHEMA_CREATE_SQL);
    validateDatabase(this.db,SCHEMA_COLUMNS);
    upgrade?.(this.db);
    retained();
    this.db.exec('PRAGMA user_version='+this.schemaVersion+';');
    return version===0;
   }).immediate();
   this.db.pragma('journal_mode = WAL');this.db.pragma('synchronous = FULL');
  }catch(error){this.db.close();throw error;}
 }

 load(scope:string,id:string){return (this.db.prepare('SELECT data FROM documents WHERE scope=? AND id=?').get(scope,id) as {data:Uint8Array}|undefined)?.data??null;}
 append(scope:string,id:string,update:Uint8Array){const check=new Y.Doc();try{Y.applyUpdate(check,update);}finally{check.destroy();}this.db.transaction(()=>{const old=this.load(scope,id);const next=old?Y.mergeUpdates([old,update]):update;this.db.prepare('INSERT INTO documents VALUES (?,?,?) ON CONFLICT(scope,id) DO UPDATE SET data=excluded.data').run(scope,id,Buffer.from(next));}).immediate();}
 preference(scope:string,key:string,value:string){return this.db.transaction(()=>{if(key==='mydoc:reading-positions:v1'){const row=this.db.prepare('SELECT value FROM preferences WHERE scope=? AND key=?').get(scope,key) as {value:string}|undefined;let old:Record<string,{updatedAt:number}>={};try{old=JSON.parse(row?.value??'{}');}catch{}const incoming=JSON.parse(value) as typeof old;for(const[k,v]of Object.entries(incoming))if(!old[k]||v.updatedAt>=old[k].updatedAt)old[k]=v;value=JSON.stringify(Object.fromEntries(Object.entries(old).sort((a,b)=>b[1].updatedAt-a[1].updatedAt).slice(0,1000)));}this.db.prepare('INSERT INTO preferences VALUES (?,?,?) ON CONFLICT(scope,key) DO UPDATE SET value=excluded.value').run(scope,key,value);return value;}).immediate();}
 close(){this.db.close();}
}

function retainRows(db:SQLiteConnection,tables:Record<string,string>){
 const retained:{table:string;columns:string;snapshot:string}[]=[];
 for(const [table,columns]of Object.entries(tables)){
  if(!db.prepare("SELECT 1 FROM main.sqlite_master WHERE type='table' AND name=?").get(table))continue;
  // Names come only from source-controlled schema contracts, never from a request.
  const snapshot='_mydoc_startup_'+table;
  db.exec(`CREATE TEMP TABLE ${snapshot} AS SELECT ${columns} FROM main.${table}`);
  retained.push({table,columns,snapshot});
 }
 return ()=>{for(const {table,columns,snapshot}of retained){
  if(db.prepare(`SELECT 1 FROM (SELECT ${columns} FROM temp.${snapshot} EXCEPT SELECT ${columns} FROM main.${table}) LIMIT 1`).get())throw new Error(`数据库升级将丢失或改写原数据：${table}，已回滚并停止启动`);
  db.exec(`DROP TABLE temp.${snapshot}`);
 }};
}

/** 在启动写事务内校验结构与完整性；失败不能继续提供存储服务。 */
export function validateDatabase(db:SQLiteConnection,tables:Record<string,string>){
 for(const [table,expected] of Object.entries(tables)){
  const columns=new Set((db.prepare(`PRAGMA table_info(${table})`).all() as {name:string}[]).map(column=>column.name));
  const missing=expected.split(',').filter(column=>!columns.has(column));
  if(missing.length)throw new Error(`数据库结构检查失败：${table} 缺少字段 ${missing.join(',')}，请检查迁移步骤`);
 }
 const check=db.pragma('quick_check') as Record<string,unknown>[];
 if(check.length!==1||Object.values(check[0])[0]!=='ok')throw new Error('数据库完整性检查失败，请恢复备份后重试');
}
