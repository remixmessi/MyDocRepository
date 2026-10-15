import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdtemp,readFile,writeFile,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {recoverReleaseManifests} from './recover-release-manifests.mjs';
async function fixture(run){
 const dir=await mkdtemp(join(tmpdir(),'mydoc-recover-release-')),out=join(dir,'output');
 const release={tag_name:'v0.1.0',draft:false,prerelease:false,published_at:'2026-10-15T01:58:29Z',assets:[]};
 try{
  for(const name of ['MyDoc-0.1.0-mac-arm64.zip','MyDoc-0.1.0-mac-x64.zip','MyDoc-0.1.0-win-x64.exe']){
   const bytes=Buffer.from(name);await writeFile(join(dir,name),bytes);release.assets.push({name,size:bytes.length,digest:'sha256:'+createHash('sha256').update(bytes).digest('hex')});
  }
  await run(release,dir,out);
 }finally{await rm(dir,{recursive:true,force:true});}
}
test('recover both manifests using the release version, exact package hashes and publish date',()=>fixture(async(release,dir,out)=>{
 const files=await recoverReleaseManifests(release,out,[dir]);assert.equal(files.length,2);
 const mac=await readFile(join(out,'latest-mac.yml'),'utf8');assert.match(mac,/version: 0.1.0/);assert.match(mac,/mac-arm64.zip/);assert.match(mac,/mac-x64.zip/);assert.match(mac,/2026-10-15T01:58:29.000Z/);
 const checksum=createHash('sha512').update('MyDoc-0.1.0-win-x64.exe').digest('base64');assert.ok((await readFile(join(out,'latest.yml'),'utf8')).includes(checksum));
}));
test('refuse mismatched binaries, absent GitHub digests, wrong versions and missing Mac ZIPs',()=>fixture(async(release,dir,out)=>{
 await assert.rejects(recoverReleaseManifests({...release,tag_name:'v0.1.1'},out,[dir]),/no supported/);
 await assert.rejects(recoverReleaseManifests({...release,assets:[{...release.assets[0],digest:null}]},out,[dir]),/checksum/);
 await assert.rejects(recoverReleaseManifests({...release,assets:[...release.assets,{name:'MyDoc-0.1.0-mac-arm64.dmg'}].filter(asset=>!asset.name.endsWith('arm64.zip'))},out,[dir]),/missing its update ZIP/);
 await writeFile(join(dir,release.assets[0].name),'new build');await assert.rejects(recoverReleaseManifests(release,out,[dir]),/differs/);
 await assert.rejects(readFile(join(out,'latest-mac.yml')),/ENOENT/);
}));
