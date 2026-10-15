import {expect,it} from 'vitest';
import {existsSync,mkdtempSync,readFileSync,rmSync,symlinkSync,writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {checkPersistentStorage} from '../packages/frontend/apps/web/server/persistent-storage';

function fixture(){
 const directory=mkdtempSync(join(tmpdir(),'mydoc-docker-storage-')),database=join(directory,'cloud.sqlite');
 const mounts=`100 99 0:1 / ${directory} rw,relatime - ext4 /dev/test rw`;
 return {directory,database,mounts,close(){rmSync(directory,{recursive:true,force:true});}};
}

it('缺少显式挂载时即使允许首次初始化也拒绝启动，不创建数据库',()=>{
 const f=fixture();try{
  expect(()=>checkPersistentStorage(f.database,f.directory,'',true)).toThrow('数据目录未挂载');
  expect(existsSync(f.database)).toBe(false);
 }finally{f.close();}
});

it('升级挂错空目录时拒绝创建空库，仅首次安装可显式初始化',()=>{
 const f=fixture();try{
  expect(()=>checkPersistentStorage(f.database,f.directory,f.mounts)).toThrow('原数据库不存在');
  expect(existsSync(f.database)).toBe(false);
  writeFileSync(join(f.directory,'server-config.json'),'{}');
  expect(()=>checkPersistentStorage(f.database,f.directory,f.mounts,true)).not.toThrow();
  expect(existsSync(f.database)).toBe(false);
 }finally{f.close();}
});

it('已有数据库不被存储检查修改，空数据库和孤立 WAL 不自动重建',()=>{
 const f=fixture();try{
  const bytes=Buffer.from('existing database bytes');writeFileSync(f.database,bytes);
  for(const initialize of [false,true])checkPersistentStorage(f.database,f.directory,f.mounts,initialize);
  expect(readFileSync(f.database)).toEqual(bytes);
  writeFileSync(f.database,'');
  expect(()=>checkPersistentStorage(f.database,f.directory,f.mounts,true)).toThrow('数据库为空');
  rmSync(f.database);writeFileSync(f.database+'-wal','recoverable WAL');
  expect(()=>checkPersistentStorage(f.database,f.directory,f.mounts,true)).toThrow('并非空目录');
  expect(readFileSync(f.database+'-wal','utf8')).toBe('recoverable WAL');
  expect(existsSync(f.database)).toBe(false);
 }finally{f.close();}
});

it('拒绝配置文件将数据库写到持久化目录外，或通过符号链接逃逸',()=>{
 const f=fixture(),outside=mkdtempSync(join(tmpdir(),'mydoc-outside-'));
 try{
  const other=join(outside,'cloud.sqlite');writeFileSync(other,'original');
  expect(()=>checkPersistentStorage(other,f.directory,f.mounts)).toThrow('数据库路径必须位于');
  symlinkSync(other,f.database);
  expect(()=>checkPersistentStorage(f.database,f.directory,f.mounts)).toThrow('符号链接');
  expect(readFileSync(other,'utf8')).toBe('original');
 }finally{f.close();rmSync(outside,{recursive:true,force:true});}
});
