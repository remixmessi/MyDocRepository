import {cloudUser} from './cloud-plugin';
import {DatabaseBackups} from '../../../common/nbstore/server/backups';
import Database from 'better-sqlite3';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import type {IncomingMessage,ServerResponse} from 'node:http';
import type {Plugin} from 'vite';
import {SQLiteProjectDatabase} from '../../../common/nbstore/server/database';
import {storageHandler} from '../../../common/nbstore/server/http';
export {storageHandler};
export const projectDatabasePath=resolve(dirname(fileURLToPath(import.meta.url)),'../../../../data/mydoc.sqlite');
export class ProjectDatabase extends SQLiteProjectDatabase {constructor(path:string){super(path,path=>new Database(path));}}

export function sqliteServer():Plugin {let database:ProjectDatabase|undefined;let backups:DatabaseBackups|undefined;const mount=(server:{middlewares:{use:(path:string,handler:(req:IncomingMessage,res:ServerResponse)=>void)=>void};httpServer?:{once:(event:string,callback:()=>void)=>void}|null})=>{database=new ProjectDatabase(process.env.MYDOC_DB_PATH??projectDatabasePath);if(process.env.MYDOC_TEST_STORAGE!=='1'){backups=new DatabaseBackups(database,path=>new Database(path));backups.start();}server.middlewares.use('/api/storage',(req,res)=>{if(!cloudUser(req)&&req.headers['x-mydoc-account']){res.writeHead(409,{'Content-Type':'application/json'});res.end(JSON.stringify({error:'账户已退出，请重新打开页面，未保存内容请先复制。'}));return;}void storageHandler(database!,process.env.MYDOC_TEST_STORAGE==='1',cloudUser(req)?undefined:backups,cloudUser(req)?'user:'+cloudUser(req)!.id:undefined)(req,res);});server.httpServer?.once('close',()=>{backups?.stop();database?.close();});};return {name:'mydoc-project-sqlite',configureServer:mount,configurePreviewServer:mount};}
