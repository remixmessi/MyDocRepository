import {randomUUID} from 'node:crypto';
import {validateDatabase,type SQLiteConnection} from '../../../../common/nbstore/server/database';

export const CLOUD_SCHEMA_VERSION=6;
/** 由公共数据库启动事务调用；版本记录与 DDL 失败时一起回滚。 */
export function upgradeCloudDatabase(db:SQLiteConnection){const columns=(table:string)=>new Set((db.prepare(`PRAGMA table_info(${table})`).all() as {name:string}[]).map(column=>column.name));
  const add=(table:string,column:string,definition:string)=>{if(!columns(table).has(column))db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`);};
   db.exec('CREATE TABLE IF NOT EXISTS cloud_migrations(version INTEGER PRIMARY KEY)');
   const version=(db.prepare('SELECT max(version) version FROM cloud_migrations').get() as {version:number|null}).version??0;
   if(version>CLOUD_SCHEMA_VERSION)throw new Error('协作数据库版本高于当前程序，请升级程序后再启动');
   if(!Number.isSafeInteger(version)||version<0)throw new Error('协作数据库版本无效');
   db.exec(`CREATE TABLE IF NOT EXISTS cloud_deleted(workspace TEXT NOT NULL,id TEXT NOT NULL,PRIMARY KEY(workspace,id));CREATE TABLE IF NOT EXISTS cloud_users(id TEXT PRIMARY KEY,email TEXT UNIQUE NOT NULL,name TEXT NOT NULL,salt TEXT NOT NULL,password TEXT NOT NULL);CREATE TABLE IF NOT EXISTS cloud_sessions(token TEXT PRIMARY KEY,user TEXT NOT NULL,expires INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS cloud_workspaces(id TEXT PRIMARY KEY,name TEXT NOT NULL,created INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS cloud_members(workspace TEXT NOT NULL,user TEXT NOT NULL,role TEXT NOT NULL,PRIMARY KEY(workspace,user));CREATE TABLE IF NOT EXISTS cloud_invites(token TEXT PRIMARY KEY,workspace TEXT NOT NULL,email TEXT NOT NULL,role TEXT NOT NULL,expires INTEGER NOT NULL);CREATE TABLE IF NOT EXISTS cloud_tokens(token TEXT PRIMARY KEY,user TEXT NOT NULL,kind TEXT NOT NULL,payload TEXT,created INTEGER NOT NULL,expires INTEGER NOT NULL);CREATE INDEX IF NOT EXISTS cloud_tokens_user ON cloud_tokens(user,kind);CREATE TABLE IF NOT EXISTS cloud_releases(id TEXT PRIMARY KEY,version TEXT NOT NULL,platform TEXT NOT NULL,filename TEXT NOT NULL,notes TEXT NOT NULL,bytes INTEGER NOT NULL,sha256 TEXT NOT NULL,created INTEGER NOT NULL,path TEXT NOT NULL);`);
   if(version<1){db.exec("CREATE TABLE IF NOT EXISTS cloud_document_permissions(workspace TEXT NOT NULL,doc TEXT NOT NULL,user TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN ('viewer','editor')),PRIMARY KEY(workspace,doc,user))");db.prepare('INSERT INTO cloud_migrations(version) VALUES (?)').run(1);}
   if(version<2){add('cloud_users','server_admin','INTEGER NOT NULL DEFAULT 0 CHECK(server_admin IN (0,1))');db.exec("CREATE TABLE IF NOT EXISTS cloud_bootstrap(key TEXT PRIMARY KEY,user TEXT NOT NULL)");db.prepare('INSERT INTO cloud_migrations(version) VALUES (?)').run(2);}
   if(version<3){db.exec("ALTER TABLE cloud_document_permissions RENAME TO cloud_document_permissions_old;CREATE TABLE cloud_document_permissions(workspace TEXT NOT NULL,doc TEXT NOT NULL,user TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN ('viewer','editor','denied')),PRIMARY KEY(workspace,doc,user));INSERT INTO cloud_document_permissions SELECT workspace,doc,user,role FROM cloud_document_permissions_old;DROP TABLE cloud_document_permissions_old;");db.prepare('INSERT INTO cloud_migrations(version) VALUES (?)').run(3);}
   if(version<4){db.exec('CREATE TABLE IF NOT EXISTS cloud_asset_owners(workspace TEXT NOT NULL,id TEXT NOT NULL,doc TEXT,user TEXT NOT NULL,PRIMARY KEY(workspace,id))');db.prepare('INSERT INTO cloud_migrations(version) VALUES (?)').run(4);}
   // 迁移记录可能来自部分升级或手工恢复，版本号较新也要修复缺失的基础结构。
   db.exec("CREATE TABLE IF NOT EXISTS cloud_document_permissions(workspace TEXT NOT NULL,doc TEXT NOT NULL,user TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN ('viewer','editor','denied')),PRIMARY KEY(workspace,doc,user));CREATE TABLE IF NOT EXISTS cloud_asset_owners(workspace TEXT NOT NULL,id TEXT NOT NULL,doc TEXT,user TEXT NOT NULL,PRIMARY KEY(workspace,id));");
   add('cloud_users','server_admin','INTEGER NOT NULL DEFAULT 0 CHECK(server_admin IN (0,1))');
   db.exec("CREATE TABLE IF NOT EXISTS cloud_bootstrap(key TEXT PRIMARY KEY,user TEXT NOT NULL)");
   add('cloud_users','avatar','BLOB');add('cloud_users','avatar_type','TEXT');add('cloud_users','verified','INTEGER NOT NULL DEFAULT 0');add('cloud_users','created','INTEGER NOT NULL DEFAULT 0');
   add('cloud_sessions','id','TEXT');add('cloud_sessions','created','INTEGER NOT NULL DEFAULT 0');add('cloud_sessions','seen','INTEGER NOT NULL DEFAULT 0');add('cloud_sessions','agent','TEXT');add('cloud_sessions','ip','TEXT');
   // 旧会话没有公开 ID，补一个随机值，避免把 token 哈希暴露给注销接口。
   for(const row of db.prepare('SELECT token FROM cloud_sessions WHERE id IS NULL').all() as {token:string}[])db.prepare('UPDATE cloud_sessions SET id=? WHERE token=?').run(randomUUID(),row.token);
   if(version<5)db.prepare('INSERT INTO cloud_migrations(version) VALUES (?)').run(5);
   add('cloud_releases','arch',"TEXT NOT NULL DEFAULT ''");
   if(version<6)db.prepare('INSERT INTO cloud_migrations(version) VALUES (?)').run(6);
   validateDatabase(db,{
    cloud_migrations:'version',cloud_deleted:'workspace,id',cloud_users:'id,email,name,salt,password,server_admin,avatar,avatar_type,verified,created',
    cloud_sessions:'token,user,expires,id,created,seen,agent,ip',cloud_workspaces:'id,name,created',cloud_members:'workspace,user,role',
    cloud_invites:'token,workspace,email,role,expires',cloud_tokens:'token,user,kind,payload,created,expires',
    cloud_releases:'id,version,platform,filename,notes,bytes,sha256,created,path,arch',cloud_bootstrap:'key,user',
    cloud_document_permissions:'workspace,doc,user,role',cloud_asset_owners:'workspace,id,doc,user',
   });
}
