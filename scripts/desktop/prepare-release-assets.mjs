import {createRequire} from 'node:module';
import {createHash} from 'node:crypto';
import {readdir,readFile,writeFile,mkdir,copyFile,stat} from 'node:fs/promises';
import {createReadStream} from 'node:fs';
import {basename,join,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const require=createRequire(import.meta.url);
// Reuse the exact YAML parser used by the installed builder.
const {load,dump}=createRequire(require.resolve('app-builder-lib/package.json'))('js-yaml');

async function digest(path,algorithm){const hash=createHash(algorithm);for await(const chunk of createReadStream(path))hash.update(chunk);return hash.digest(algorithm==='sha512'?'base64':'hex');}

/** Merge per-runner manifests before uploading: overwriting latest-mac.yml loses one Mac architecture. */
export async function prepareReleaseAssets(input,output,version){
 if(!/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(version))throw new Error('Invalid release version');
 await mkdir(output,{recursive:true});
 const manifests=new Map(),assets=new Map();
 for(const entry of await readdir(input,{withFileTypes:true})){
  if(!entry.isDirectory())continue;
  for(const name of await readdir(join(input,entry.name))){
   const path=join(input,entry.name,name);
   if(name==='latest-mac.yml'||name==='latest.yml'){
    const info=load(await readFile(path,'utf8'));
    if(info?.version!==version||!Array.isArray(info.files)||!info.files.length)throw new Error('Release manifest version or files invalid: '+name);
    manifests.set(name,[...(manifests.get(name)??[]),info]);
   }else if(/\.(dmg|zip|exe|blockmap)$/.test(name)){
    const checksum=await digest(path,'sha512'),previous=assets.get(name);
    if(previous&&previous.sha512!==checksum)throw new Error('Conflicting release asset: '+name);
    assets.set(name,{sha512:checksum,size:(await stat(path)).size});await copyFile(path,join(output,name));
   }
  }
 }
 for(const name of [`MyDoc-${version}-mac-arm64.zip`,`MyDoc-${version}-mac-x64.zip`,`MyDoc-${version}-win-x64.exe`])if(!assets.has(name))throw new Error('Missing update package: '+name);
 for(const name of ['latest-mac.yml','latest.yml']){
  const infos=manifests.get(name);if(!infos?.length)throw new Error('Missing update manifest: '+name);
  const files=new Map();
  for(const info of infos)for(const file of info.files){
   if(typeof file.url!=='string'||basename(file.url)!==file.url||file.url.includes('\\'))throw new Error('Invalid update asset path');
   const asset=assets.get(file.url);
   if(!asset||asset.sha512!==file.sha512||file.size!==undefined&&file.size!==asset.size)throw new Error('Update asset checksum/size mismatch: '+file.url);
   files.set(file.url,file);
  }
  const merged=[...files.values()].sort((a,b)=>Number(b.url.endsWith('.zip'))-Number(a.url.endsWith('.zip'))||a.url.localeCompare(b.url));
  const required=name==='latest-mac.yml'?[`MyDoc-${version}-mac-arm64.zip`,`MyDoc-${version}-mac-x64.zip`]:[`MyDoc-${version}-win-x64.exe`];
  if(required.some(file=>!files.has(file)))throw new Error('Manifest is missing update architectures: '+name);
  await writeFile(join(output,name),dump({...infos[0],files:merged,path:merged[0].url,sha512:merged[0].sha512}));
 }
 const checksums=[];
 for(const name of (await readdir(output)).filter(name=>name!=='SHA256SUMS.txt').sort())checksums.push(await digest(join(output,name),'sha256')+'  '+name);
 await writeFile(join(output,'SHA256SUMS.txt'),checksums.join('\n')+'\n');
}
if(process.argv[1]&&resolve(process.argv[1])===fileURLToPath(import.meta.url))await prepareReleaseAssets(...process.argv.slice(2));
