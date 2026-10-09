import {existsSync} from 'node:fs';
import {readFile,writeFile,mkdir} from 'node:fs/promises';import {spawnSync} from 'node:child_process';import {fileURLToPath} from 'node:url';import {dirname,resolve} from 'node:path';import {releaseConfig,validateRelease} from './release-config.mjs';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'../..'),desktop=resolve(root,'packages/frontend/apps/desktop');const target=process.argv[2];
try{validateRelease(target,process.env,process.platform);}catch(error){console.error(error.message);process.exit(1);}
if(process.argv.includes('--check')){console.log('签名／公证配置检查通过（尚未构建或发布）');process.exit(0);}
const base=JSON.parse(await readFile(resolve(desktop,'electron-builder.json'),'utf8'));await mkdir(resolve(desktop,'dist'),{recursive:true});
const env={...process.env,MYDOC_SIGNED_RELEASE:'1',CSC_IDENTITY_AUTO_DISCOVERY:'true'};
// Keep the repository-pinned Corepack Yarn.
function yarn(args){const cli=process.env.COREPACK_ROOT?resolve(process.env.COREPACK_ROOT,'dist/yarn.js'):process.env.npm_execpath;if(!cli||!existsSync(cli)||!/[.](?:cjs|mjs|js)$/.test(cli))throw new Error('请通过 corepack yarn desktop:release:mac 或 desktop:release:win 运行');const result=spawnSync(process.execPath,[cli,...args],{cwd:root,env,stdio:'inherit'});if(result.error)throw result.error;if(result.status!==0)process.exit(result.status??1);}
yarn(['desktop:build']);const config=resolve(desktop,'dist/release-builder.json');await writeFile(config,JSON.stringify(releaseConfig(base,env)));yarn(['workspace','@mydoc/desktop','exec','electron-builder','--config',config,target==='mac'?'--mac':'--win',...(target==='mac'?['--arm64','--x64']:[]),'--publish','never']);
