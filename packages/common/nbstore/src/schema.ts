/**
 * SQLite 结构版本与离线队列／草稿表定义。浏览器端与服务端共用，不依赖 node 模块。
 *
 * 版本 1：documents、assets、preferences 三张业务表。
 * 版本 2：新增 pending_updates（待提交 Yjs 更新队列）与 drafts（草稿恢复条目）。
 * 版本 3：将本地账号和会话表纳入启动迁移（兼容此前请求时创建的表）。
 * 版本 4：附件文件索引；启动将历史 BLOB 校验迁移到数据库同级 attachments/。
 *
 * 业务表参与完整备份恢复；队列与草稿属于本机未提交状态，恢复时保留不被覆盖，
 * 以 generation 列标记写入时的数据库代际，供恢复后识别旧代际条目。
 */
export const SCHEMA_VERSION = 4;
/** 完整备份覆盖的业务表及其列顺序。恢复只替换这些表。 */
export const BUSINESS_TABLES: [string, string][] = [
  ['documents', 'scope,id,data'],
  ['assets', 'scope,workspace,id,name,mime,data'],
  ['preferences', 'scope,key,value'],
];
/** 本机未提交状态表。随快照文件存在，但不参与恢复替换。 */
export const LOCAL_STATE_TABLES = ['pending_updates', 'drafts'];
/** 整库快照包含账号，但恢复业务数据时不替换账号与登录会话。 */
export const LOCAL_ACCOUNT_TABLES = ['local_accounts', 'local_account_sessions'];

/**
 * 隔离键：scope 已编码账户（`main` 或 `user:<账户 ID>`），server 编码目标服务器，
 * workspace 编码工作空间，doc 为 `workspace:<id>` 或 `doc:<id>`。
 * 同一账户连接多台服务器、或同一服务器下多个空间，条目互不混用。
 */
export const SCHEMA_MIGRATIONS: Record<number, string> = {
  1: `CREATE TABLE IF NOT EXISTS documents(scope TEXT NOT NULL,id TEXT NOT NULL,data BLOB NOT NULL,PRIMARY KEY(scope,id));
CREATE TABLE IF NOT EXISTS assets(scope TEXT NOT NULL,workspace TEXT NOT NULL,id TEXT NOT NULL,name TEXT NOT NULL,mime TEXT NOT NULL,data BLOB NOT NULL,PRIMARY KEY(scope,workspace,id));
CREATE TABLE IF NOT EXISTS preferences(scope TEXT NOT NULL,key TEXT NOT NULL,value TEXT NOT NULL,PRIMARY KEY(scope,key));`,
  2: `CREATE TABLE IF NOT EXISTS pending_updates(seq INTEGER PRIMARY KEY,scope TEXT NOT NULL,server TEXT NOT NULL,workspace TEXT NOT NULL,doc TEXT NOT NULL,data BLOB NOT NULL,created INTEGER NOT NULL,attempts INTEGER NOT NULL DEFAULT 0,retryAt INTEGER NOT NULL DEFAULT 0,error TEXT,generation TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS pending_updates_target ON pending_updates(scope,server,workspace,doc,seq);
CREATE TABLE IF NOT EXISTS drafts(id TEXT PRIMARY KEY,scope TEXT NOT NULL,server TEXT NOT NULL,workspace TEXT NOT NULL,doc TEXT NOT NULL,title TEXT NOT NULL,summary TEXT NOT NULL,data BLOB NOT NULL,missingAssets INTEGER NOT NULL DEFAULT 0,created INTEGER NOT NULL,generation TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS drafts_target ON drafts(scope,server,workspace,created);`,
  3: `CREATE TABLE IF NOT EXISTS local_accounts(id TEXT PRIMARY KEY,email TEXT UNIQUE NOT NULL,name TEXT NOT NULL,salt TEXT NOT NULL,password TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS local_account_sessions(token TEXT PRIMARY KEY,user TEXT NOT NULL,expires INTEGER NOT NULL);`,
  4: `CREATE TABLE IF NOT EXISTS asset_files(scope TEXT NOT NULL,workspace TEXT NOT NULL,id TEXT NOT NULL,path TEXT NOT NULL,size INTEGER NOT NULL,sha256 TEXT NOT NULL,PRIMARY KEY(scope,workspace,id));`,
};

/** 每次启动检查的列契约；新增字段时同时追加迁移并更新这里。 */
export const SCHEMA_COLUMNS: Record<string, string> = {
 asset_files:'scope,workspace,id,path,size,sha256',
 documents:'scope,id,data', assets:'scope,workspace,id,name,mime,data', preferences:'scope,key,value',
 pending_updates:'seq,scope,server,workspace,doc,data,created,attempts,retryAt,error,generation',
 drafts:'id,scope,server,workspace,doc,title,summary,data,missingAssets,created,generation',
 local_accounts:'id,email,name,salt,password', local_account_sessions:'token,user,expires',
};

/** 仅用于补建当前表/索引，不重放未来的 ALTER 或数据迁移。 */
export const SCHEMA_CREATE_SQL = [SCHEMA_MIGRATIONS[1],SCHEMA_MIGRATIONS[2],SCHEMA_MIGRATIONS[3],SCHEMA_MIGRATIONS[4]].join('\n');
