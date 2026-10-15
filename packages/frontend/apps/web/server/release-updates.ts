import {createHash} from 'node:crypto';
import {createReadStream} from 'node:fs';
import {compareReleaseVersions,validReleaseVersion} from '../../../../common/nbstore/src/releases';
export interface UpdateReleaseRow {id:string;version:string;platform:string;arch:string;filename:string;notes:string;bytes:number;sha256:string;created:number;path:string;}
export class ReleaseUpdateError extends Error {constructor(message:string,readonly status:number){super(message);}}
export function selectUpdateRelease(rows:UpdateReleaseRow[],platform:string,arch:string){
 const stable=rows.filter(row=>row.platform===platform&&row.arch===arch&&validReleaseVersion(row.version)&&!row.version.split('+')[0].includes('-')).sort((a,b)=>compareReleaseVersions(b.version,a.version)||b.created-a.created);
 if(!stable.length)throw new ReleaseUpdateError('暂无匹配的正式更新包，请上传并明确标注系统与架构',404);
 const latest=stable[0].version,extension=platform==='darwin'?'.zip':'.exe';
 const selected=stable.find(row=>compareReleaseVersions(row.version,latest)===0&&row.filename.toLowerCase().endsWith(extension));
 if(!selected)throw new ReleaseUpdateError('最新正式版本缺少 '+(platform==='darwin'?'macOS ZIP':'Windows EXE')+' 更新包',409);
 return selected;
}
/** Hash immutable release files on demand; retain no file bytes or extra database state. */
export class ReleaseUpdateManifests {
 private checksums=new Map<string,Promise<string>>();
 async create(row:UpdateReleaseRow){
  const key=row.id+':'+row.sha256;
  let checksum=this.checksums.get(key);
  if(!checksum){
   checksum=(async()=>{const sha512=createHash('sha512'),sha256=createHash('sha256');let size=0;for await(const chunk of createReadStream(row.path)){sha512.update(chunk);sha256.update(chunk);size+=chunk.length;}if(size!==row.bytes||sha256.digest('hex')!==row.sha256)throw new ReleaseUpdateError('更新包完整性校验失败，请管理员重新上传',409);return sha512.digest('base64');})();
   this.checksums.set(key,checksum);if(this.checksums.size>100)this.checksums.delete(this.checksums.keys().next().value!);
   void checksum.catch(()=>this.checksums.delete(key));
  }
  const sha512=await checksum,url=row.id+(row.platform==='darwin'?'.zip':'.exe');
  return {version:row.version,files:[{url,sha512,size:row.bytes}],path:url,sha512,releaseDate:new Date(row.created).toISOString(),releaseNotes:row.notes,mydocPlatform:row.platform,mydocArch:row.arch};
 }
}
