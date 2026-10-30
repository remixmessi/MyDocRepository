/** Public schema 5: private Agent records; no credentials or document content ownership. */
export const AGENT_COLUMNS:Record<string,string>={
 agent_sessions:'scope,id,workspace,title,created,updated',
 agent_runs:'scope,id,session,request_key,input,status,generation,db_generation,created',
 agent_events:'scope,session,cursor,run,run_seq,payload',
 agent_messages:'scope,session,cursor,run,payload',
 agent_tool_calls:'scope,run,id,payload,result',
 agent_changesets:'scope,id,session,resource,payload,status',
 agent_approvals:'scope,id,run,hash,expires,decision',
};
export const AGENT_CREATE_SQL=`
CREATE TABLE IF NOT EXISTS agent_sessions(scope TEXT NOT NULL,id TEXT NOT NULL,workspace TEXT NOT NULL,title TEXT NOT NULL,created INTEGER NOT NULL,updated INTEGER NOT NULL,PRIMARY KEY(scope,id));
CREATE INDEX IF NOT EXISTS agent_sessions_recent ON agent_sessions(scope,workspace,updated);
CREATE TABLE IF NOT EXISTS agent_runs(scope TEXT NOT NULL,id TEXT NOT NULL,session TEXT NOT NULL,request_key TEXT NOT NULL,input TEXT NOT NULL,status TEXT NOT NULL,generation INTEGER NOT NULL,db_generation TEXT NOT NULL,created INTEGER NOT NULL,PRIMARY KEY(scope,id),UNIQUE(scope,session,request_key));
CREATE INDEX IF NOT EXISTS agent_runs_queue ON agent_runs(scope,session,status,created);
CREATE TABLE IF NOT EXISTS agent_events(scope TEXT NOT NULL,session TEXT NOT NULL,cursor INTEGER NOT NULL,run TEXT NOT NULL,run_seq INTEGER NOT NULL,payload TEXT NOT NULL,PRIMARY KEY(scope,session,cursor),UNIQUE(scope,run,run_seq));
CREATE TABLE IF NOT EXISTS agent_messages(scope TEXT NOT NULL,session TEXT NOT NULL,cursor INTEGER NOT NULL,run TEXT NOT NULL,payload TEXT NOT NULL,PRIMARY KEY(scope,session,cursor));
CREATE TABLE IF NOT EXISTS agent_tool_calls(scope TEXT NOT NULL,run TEXT NOT NULL,id TEXT NOT NULL,payload TEXT NOT NULL,result TEXT,PRIMARY KEY(scope,run,id));
CREATE TABLE IF NOT EXISTS agent_changesets(scope TEXT NOT NULL,id TEXT NOT NULL,session TEXT NOT NULL,resource TEXT NOT NULL,payload TEXT NOT NULL,status TEXT NOT NULL,PRIMARY KEY(scope,id));
CREATE TABLE IF NOT EXISTS agent_approvals(scope TEXT NOT NULL,id TEXT NOT NULL,run TEXT NOT NULL,hash TEXT NOT NULL,expires INTEGER NOT NULL,decision TEXT,PRIMARY KEY(scope,id));
`;
