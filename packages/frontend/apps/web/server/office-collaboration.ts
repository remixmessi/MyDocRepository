import {randomUUID} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import type {CommentNotice} from './office-comments';

/** Consecutive edits by one user within this window extend the same activity row. */
export const ACTIVITY_SESSION_MS=10*60*1000;
export const ACTIVITY_LIMIT_PER_RESOURCE=200;
export const NOTIFICATION_LIMIT_PER_USER=500;

export type OfficeActivity={user:string;userName:string;started:number;last:number;edits:number};
export type CloudNotification={id:string;workspace:string;workspaceName:string;resource:string;resourceName:string;kind:'mention'|'reply';actor:string;actorName:string;thread:string;excerpt:string;created:number;read:boolean};

/** Records one accepted edit. Rows are coarse sessions, not an operation log, so storage stays bounded. */
export function recordOfficeEdit(database:SQLiteProjectDatabase,workspace:string,resource:string,user:{id:string;name:string},now=Date.now()){
 const db=database.db;
 db.transaction(()=>{
  if(!db.prepare('SELECT 1 FROM cloud_office_resources WHERE workspace=? AND id=? AND deleted=0').get(workspace,resource))return;
  const open=db.prepare('SELECT started FROM cloud_office_activity WHERE workspace=? AND resource=? AND user=? AND last>=? ORDER BY last DESC LIMIT 1').get(workspace,resource,user.id,now-ACTIVITY_SESSION_MS) as {started:number}|undefined;
  if(open)db.prepare('UPDATE cloud_office_activity SET last=max(last,?),edits=edits+1,user_name=? WHERE workspace=? AND resource=? AND user=? AND started=?').run(now,user.name,workspace,resource,user.id,open.started);
  else{
   db.prepare('INSERT INTO cloud_office_activity(workspace,resource,user,user_name,started,last,edits) VALUES (?,?,?,?,?,?,1) ON CONFLICT(workspace,resource,user,started) DO UPDATE SET last=excluded.last,edits=edits+1').run(workspace,resource,user.id,user.name,now,now);
   db.prepare('DELETE FROM cloud_office_activity WHERE workspace=? AND resource=? AND rowid NOT IN (SELECT rowid FROM cloud_office_activity WHERE workspace=? AND resource=? ORDER BY last DESC LIMIT ?)').run(workspace,resource,workspace,resource,ACTIVITY_LIMIT_PER_RESOURCE);
  }
 }).immediate();
}

export function listOfficeActivity(database:SQLiteProjectDatabase,workspace:string,resource:string,limit=50):OfficeActivity[]{
 return database.db.prepare('SELECT user,user_name AS userName,started,last,edits FROM cloud_office_activity WHERE workspace=? AND resource=? ORDER BY last DESC LIMIT ?').all(workspace,resource,Math.max(1,Math.min(limit,ACTIVITY_LIMIT_PER_RESOURCE))) as OfficeActivity[];
}

/** Must run inside the comment transaction so a retried operation never duplicates notifications. */
export function insertNotifications(database:SQLiteProjectDatabase,workspace:string,resource:string,actor:{id:string;name:string},notices:CommentNotice[],now=Date.now()){
 const db=database.db,users=new Set<string>();
 for(const notice of notices){
  db.prepare('INSERT INTO cloud_notifications(id,user,workspace,resource,kind,actor,actor_name,thread,excerpt,created,read) VALUES (?,?,?,?,?,?,?,?,?,?,0)').run(randomUUID(),notice.user,workspace,resource,notice.kind,actor.id,actor.name,notice.thread,notice.excerpt,now);
  users.add(notice.user);
 }
 for(const user of users)db.prepare('DELETE FROM cloud_notifications WHERE user=? AND id NOT IN (SELECT id FROM cloud_notifications WHERE user=? ORDER BY created DESC,id DESC LIMIT ?)').run(user,user,NOTIFICATION_LIMIT_PER_USER);
 return [...users];
}

type NotificationRow={id:string;workspace:string;workspaceName:string;resource:string;resourceName:string;kind:'mention'|'reply';actor:string;actorName:string;thread:string;excerpt:string;created:number;read:number};

/** Lists the user's notifications; `visible` re-checks membership and file access so revoked files disappear. */
export function listNotifications(database:SQLiteProjectDatabase,user:string,visible:(workspace:string,resource:string)=>boolean,limit=50):{items:CloudNotification[];unread:number}{
 const rows=database.db.prepare('SELECT n.id,n.workspace,w.name AS workspaceName,n.resource,r.name AS resourceName,n.kind,n.actor,n.actor_name AS actorName,n.thread,n.excerpt,n.created,n.read FROM cloud_notifications n JOIN cloud_workspaces w ON w.id=n.workspace JOIN cloud_office_resources r ON r.workspace=n.workspace AND r.id=n.resource AND r.deleted=0 WHERE n.user=? ORDER BY n.created DESC,n.id DESC LIMIT ?').all(user,NOTIFICATION_LIMIT_PER_USER) as NotificationRow[];
 const allowed=new Map<string,boolean>(),items:CloudNotification[]=[];let unread=0;
 for(const row of rows){
  const key=row.workspace+'\n'+row.resource;if(!allowed.has(key))allowed.set(key,visible(row.workspace,row.resource));if(!allowed.get(key))continue;
  if(!row.read)unread++;if(items.length<limit)items.push({...row,read:!!row.read});
 }
 return {items,unread};
}

/** Marks ids (or everything when `ids` is undefined) as read for this user only. */
export function markNotificationsRead(database:SQLiteProjectDatabase,user:string,ids?:string[]){
 const db=database.db;
 return db.transaction(():number=>{
  if(!ids)return (db.prepare('UPDATE cloud_notifications SET read=1 WHERE user=? AND read=0').run(user) as {changes:number}).changes;
  let changed=0;for(const id of ids)changed+=(db.prepare('UPDATE cloud_notifications SET read=1 WHERE user=? AND id=? AND read=0').run(user,id) as {changes:number}).changes;return changed;
 }).immediate();
}
