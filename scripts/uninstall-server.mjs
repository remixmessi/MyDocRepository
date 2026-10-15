import {lstatSync,openSync,readSync,closeSync,existsSync,rmSync} from 'node:fs';
import {resolve,join,parse} from 'node:path';
import {homedir} from 'node:os';
// The deployment's process supervisor must stop the service before invoking this.
const args=process.argv.slice(2),position=args.indexOf('--data-dir');
if(!args.includes('--confirm-delete-data')||!args.includes('--service-stopped')||position<0||!args[position+1]||args[position+1].startsWith('--')){
 console.error('停止服务后运行：node scripts/uninstall-server.mjs --data-dir /实际数据目录 --service-stopped --confirm-delete-data');process.exit(2);
}
const directory=resolve(args[position+1]);
if([parse(directory).root,homedir(),process.cwd()].includes(directory)||lstatSync(directory).isSymbolicLink())throw Error('拒绝清理根目录、主目录、当前目录或符号链接');
const validDatabase=path=>{if(lstatSync(path).isSymbolicLink()||!lstatSync(path).isFile())return false;const fd=openSync(path,'r');try{const header=Buffer.alloc(16);readSync(fd,header,0,16,0);return header.toString()==='SQLite format 3\0';}finally{closeSync(fd);}};
const filePosition=args.indexOf('--database-file'),filename=filePosition<0?undefined:args[filePosition+1];
if(filePosition>=0&&(!filename||parse(filename).base!==filename||!filename.endsWith('.sqlite')))throw Error('数据库文件名必须是同目录的 .sqlite 文件名');
const databases=(filename?[filename]:['cloud.sqlite','mydoc.sqlite']).filter(file=>existsSync(join(directory,file)));
if(!databases.length||databases.some(file=>!validDatabase(join(directory,file))))throw Error('未找到有效 MyDoc 数据库，未删除文件');
for(const file of [...databases.flatMap(file=>[file,file+'-wal',file+'-shm',file+'-journal']),'attachments','backups','upgrade-backups','upgrade-inputs','releases','backup-settings.json','backup-settings.json.tmp','server-config.json'])rmSync(join(directory,file),{recursive:true,force:true});
console.log('已删除该服务的数据、附件、内部备份和配置：'+directory);
