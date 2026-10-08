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

export function sqliteServer():Plugin {let database:ProjectDatabase|undefined;const mount=(server:{middlewares:{use:(path:string,handler:(req:IncomingMessage,res:ServerResponse)=>void)=>void};httpServer?:{once:(event:string,callback:()=>void)=>void}|null})=>{database=new ProjectDatabase(process.env.MYDOC_DB_PATH??projectDatabasePath);server.middlewares.use('/api/storage',(req,res)=>void storageHandler(database!,process.env.MYDOC_TEST_STORAGE==='1')(req,res));server.httpServer?.once('close',()=>database?.close());};return {name:'mydoc-project-sqlite',configureServer:mount,configurePreviewServer:mount};}
