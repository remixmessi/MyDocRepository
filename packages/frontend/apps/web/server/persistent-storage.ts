import {existsSync,readFileSync,readdirSync,realpathSync,statSync} from 'node:fs';
import {dirname,isAbsolute,relative,resolve,sep} from 'node:path';

/** Validate storage before opening SQLite: an upgrade must never create a replacement database. */
export function checkPersistentStorage(databasePath:string,dataDirectory:string,mountInfo:string,initialize=false){
 const directory=resolve(dataDirectory),database=resolve(databasePath);
 const mounted=mountInfo.split('\n').some(line=>{
  const target=line.split(' ')[4];
  return target?.replace(/\\([0-7]{3})/g,(_,octal:string)=>String.fromCharCode(parseInt(octal,8)))===directory;
 });
 if(!mounted)throw new Error(`数据目录未挂载：${directory}。请显式挂载原服务器目录或原 Docker 卷，拒绝在容器内创建空库。`);
 const inside=(base:string,path:string,allowRoot=false)=>{const child=relative(base,path);return (allowRoot||child!=='')&&!child.startsWith('..'+sep)&&child!=='..'&&!isAbsolute(child);};
 const realDirectory=realpathSync(directory);
 if(!inside(directory,database)||!inside(realDirectory,realpathSync(dirname(database)),true))throw new Error(`数据库路径必须位于持久化目录 ${directory} 内：${database}。请检查 cloudDbPath 配置。`);
 if(existsSync(database)){
  if(!inside(realDirectory,realpathSync(database)))throw new Error('数据库符号链接指向持久化目录外，拒绝启动');
  if(!statSync(database).isFile()||statSync(database).size===0)throw new Error(`数据库为空或不是普通文件：${database}。拒绝重建，请检查原数据卷和备份。`);
  return;
 }
 if(!initialize)throw new Error(`原数据库不存在：${database}。可能挂载了新目录或新卷，已停止启动以避免创建空库。仅首次安装可临时设置 MYDOC_INIT_DATABASE=1。`);
 // An orphan WAL or a populated directory may still contain recoverable data.
 if(readdirSync(directory).some(name=>name!=='server-config.json'))throw new Error(`数据目录 ${directory} 并非空目录但缺少数据库，拒绝初始化。请先查找原数据库、WAL 和备份。`);
}

export function assertPersistentStorage(databasePath:string,env:NodeJS.ProcessEnv=process.env){
 if(env.MYDOC_CONTAINER!=='1')return;
 checkPersistentStorage(databasePath,'/app/data',readFileSync('/proc/self/mountinfo','utf8'),env.MYDOC_INIT_DATABASE==='1');
}
