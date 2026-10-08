import {build} from 'esbuild';
import {cp,mkdir,rm} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {resolve,dirname} from 'node:path';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..'),desktop=resolve(root,'packages/frontend/apps/desktop'),output=resolve(desktop,'dist');
await rm(output,{recursive:true,force:true});await mkdir(output,{recursive:true});
for(const entry of ['main','preload'])await build({entryPoints:[resolve(desktop,'src',entry+'.ts')],outfile:resolve(output,entry+'.cjs'),bundle:true,platform:'node',format:'cjs',target:'node24',external:['electron'],sourcemap:true});
await cp(resolve(root,'packages/frontend/apps/web/dist'),resolve(output,'renderer'),{recursive:true});
console.log('Desktop main, preload and renderer built.');
