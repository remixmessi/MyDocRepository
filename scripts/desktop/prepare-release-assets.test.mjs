import {test} from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,readFile,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {createHash} from 'node:crypto';
import {prepareReleaseAssets} from './prepare-release-assets.mjs';
async function fixture(run){
 const dir=await mkdtemp(join(tmpdir(),'mydoc-update-manifests-')),input=join(dir,'artifacts'),output=join(dir,'assets');await mkdir(input);
 try{
  for(const [os,arch,ext] of [['mac','arm64','zip'],['mac','x64','zip'],['win','x64','exe']]){
   const target=join(input,os+'-'+arch);await mkdir(target);const name=`MyDoc-1.2.3-${os}-${arch}.${ext}`,bytes=Buffer.from(name);
   await writeFile(join(target,name),bytes);
   await writeFile(join(target,os==='mac'?'latest-mac.yml':'latest.yml'),JSON.stringify({version:'1.2.3',files:[{url:name,sha512:createHash('sha512').update(bytes).digest('base64'),size:bytes.length}]}));
  }
  await run(input,output);
 }finally{await rm(dir,{recursive:true,force:true});}
}
test('merge Mac architectures, preserve Windows manifest, verify package digests and produce checksums',()=>fixture(async(input,output)=>{
 await prepareReleaseAssets(input,output,'1.2.3');const mac=await readFile(join(output,'latest-mac.yml'),'utf8');assert.match(mac,/mac-arm64.zip/);assert.match(mac,/mac-x64.zip/);assert.match(await readFile(join(output,'latest.yml'),'utf8'),/win-x64.exe/);assert.match(await readFile(join(output,'SHA256SUMS.txt'),'utf8'),/latest-mac.yml/);
}));
test('reject mismatched version and tampered package before upload',()=>fixture(async(input,output)=>{
 await assert.rejects(prepareReleaseAssets(input,output,'1.2.4'),/version/);
 await writeFile(join(input,'mac-arm64','MyDoc-1.2.3-mac-arm64.zip'),'corrupted');await assert.rejects(prepareReleaseAssets(input,output,'1.2.3'),/checksum/);
}));
