import {createRequire} from 'node:module';
import {createHash} from 'node:crypto';
import {createReadStream,existsSync} from 'node:fs';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import {join,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const require=createRequire(import.meta.url);
const {dump}=createRequire(require.resolve('app-builder-lib/package.json'))('js-yaml');

/** Recover manifests from the exact published binaries, never a new build or the working tree version. No network writes. */
export async function recoverReleaseManifests(release,output,directories){
 const version=typeof release?.tag_name==='string'?release.tag_name.replace(/^v/,''):'';
 if(!/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(version)||release.draft!==false||release.prerelease!==false||!Array.isArray(release.assets)||!Number.isFinite(Date.parse(release.published_at)))throw new Error('Expected a published stable GitHub Release');
 const expected=new Map();
 for(const asset of release.assets){
  const prefix=`MyDoc-${version}-`;
  if(typeof asset.name!=='string'||!asset.name.startsWith(prefix))continue;
  const target=asset.name.slice(prefix.length);
  if(/^(mac-(arm64|x64)\.zip|win-(arm64|x64)\.exe)$/.test(target))expected.set(asset.name,{asset,channel:target.startsWith('mac-')?'latest-mac.yml':'latest.yml'});
 }
 if(!expected.size)throw new Error('Release has no supported update ZIP/EXE packages');
 for(const asset of release.assets){
  if(new RegExp('^MyDoc-'+version.replaceAll('.','\\.')+'-mac-(arm64|x64)\\.dmg$').test(asset.name)&&!expected.has(asset.name.replace(/\.dmg$/,'.zip')))throw new Error('macOS Release is missing its update ZIP: '+asset.name);
 }
 const groups=new Map();
 for(const [name,{asset,channel}] of expected){
  if(!/^sha256:[a-f0-9]{64}$/.test(asset.digest??'')||!Number.isSafeInteger(asset.size)||asset.size<=0)throw new Error('Missing authoritative GitHub checksum/size: '+name);
  const path=directories.map(dir=>join(dir,name)).find(path=>existsSync(path));if(!path)throw new Error('Download the published package first: '+name);
  const sha256=createHash('sha256'),sha512=createHash('sha512');let size=0;
  for await(const chunk of createReadStream(path)){sha256.update(chunk);sha512.update(chunk);size+=chunk.length;}
  if('sha256:'+sha256.digest('hex')!==asset.digest||size!==asset.size)throw new Error('Local package differs from the published GitHub asset: '+name);
  groups.set(channel,[...(groups.get(channel)??[]),{url:name,sha512:sha512.digest('base64'),size}]);
 }
 await mkdir(output,{recursive:true});
 const result=[];
 for(const [channel,files] of groups){
  files.sort((a,b)=>a.url.localeCompare(b.url));const path=join(output,channel);
  await writeFile(path,dump({version,files,path:files[0].url,sha512:files[0].sha512,releaseDate:new Date(release.published_at).toISOString()}));result.push(path);
 }
 return result;
}
if(process.argv[1]&&resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 const [metadata,output,...directories]=process.argv.slice(2);
 if(!metadata||!output||!directories.length)throw new Error('Usage: node recover-release-manifests.mjs release.json output-dir package-dir [...]');
 const files=await recoverReleaseManifests(JSON.parse(await readFile(metadata,'utf8')),output,directories);
 console.log('Verified published binaries; generated '+files.join(', '));
}
