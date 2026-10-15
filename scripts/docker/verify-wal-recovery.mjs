import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {randomUUID} from 'node:crypto';
import {setTimeout as delay} from 'node:timers/promises';

const original=process.env.MYDOC_OLD_IMAGE||'mydoc-web:latest';
const fixed=process.env.MYDOC_UPGRADE_IMAGE||'mydoc-web:2026.10.15-1';
const name='mydoc-wal-'+randomUUID().slice(0,8),volume=name+'-data',password=randomUUID();
function run(args){const result=spawnSync('docker',args,{encoding:'utf8',maxBuffer:2e6});if(result.error)throw result.error;return result;}
function docker(args){const result=run(args);assert.equal(result.status,0,result.stderr||result.stdout);return result.stdout.trim();}
const mounted=(image,script)=>docker(['run','--rm','--network','none','--entrypoint','node','-v',volume+':/app/data',image,'-e',script]);
const fingerprints=()=>mounted(fixed,`const fs=require('node:fs'),crypto=require('node:crypto'),root='/app/data/cloud.sqlite';console.log(JSON.stringify(Object.fromEntries(['','-wal','-shm','-journal'].filter(s=>fs.existsSync(root+s)).map(s=>[s,crypto.createHash('sha256').update(fs.readFileSync(root+s)).digest('hex')]))));`);
try{
 docker(['run','-d','--name',name,'--platform','linux/amd64','-p','127.0.0.1::8080','-v',volume+':/app/data','-e','MYDOC_ADMIN_EMAIL=original@example.com','-e','MYDOC_ADMIN_PASSWORD='+password,original]);
 const root='http://'+docker(['port',name,'8080/tcp']);let ready=false;
 for(let i=0;i<90;i++){try{if((await fetch(root+'/api/hosting',{signal:AbortSignal.timeout(1000)})).ok){ready=true;break;}}catch{}await delay(300);}assert.ok(ready,docker(['logs',name]));
 const login=await fetch(root+'/api/cloud/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'original@example.com',password})});assert.ok(login.ok);
 const workspace=await fetch(root+'/api/cloud/workspaces',{method:'POST',headers:{Cookie:login.headers.get('set-cookie').split(';')[0],'Content-Type':'application/json'},body:JSON.stringify({name:'original team'})});assert.ok(workspace.ok);
 docker(['stop',name]);docker(['rm',name]);
 // Simulate restoring the correct old main file into a directory containing a
 // WAL from a database in which a different administrator was initialized.
 mounted(original,`const fs=require('node:fs'),Database=require('better-sqlite3'),main='/app/data/cloud.sqlite',other='/app/data/other.sqlite';fs.copyFileSync(main,other);const db=new Database(other);db.pragma('wal_autocheckpoint=0');db.transaction(()=>{db.exec('DELETE FROM cloud_members;DELETE FROM cloud_workspaces;DELETE FROM cloud_users;DELETE FROM cloud_bootstrap');db.prepare('INSERT INTO cloud_users(id,email,name,salt,password) VALUES(?,?,?,?,?)').run('replacement-admin','replacement@example.com','replacement','salt','password');})();fs.copyFileSync(other+'-wal',main+'-wal');db.close();fs.unlinkSync(other);`);
 const before=fingerprints();
 const failed=run(['run','--rm','--network','none','-v',volume+':/app/data',fixed]);
 assert.notEqual(failed.status,0);assert.match(failed.stderr+failed.stdout,/WAL 应用后已有账号消失/);
 assert.equal(fingerprints(),before,'the original main file and WAL must remain byte-for-byte unchanged');
 mounted(fixed,`const fs=require('node:fs'),root='/app/data/upgrade-inputs',entries=fs.readdirSync(root);if(!entries.some(x=>x.startsWith('input-')))throw Error('missing raw input backup');`);
 console.log('PASS: final image refuses a foreign WAL that removes old accounts before opening the original database');
 console.log('PASS: original main/WAL bytes are unchanged and their raw copies remain available');
}finally{
 spawnSync('docker',['rm','-f',name],{stdio:'ignore'});
 spawnSync('docker',['volume','rm',volume],{stdio:'ignore'});
}
