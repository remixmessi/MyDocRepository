import {fileURLToPath} from 'node:url';
import {statSync} from 'node:fs';
import {resolve} from 'node:path';
import {build} from 'esbuild';
let cached:{inputs:[string,number][];text:string}|undefined;
/** Used only by Vite/tests. Production embeds the parser bundle at build time.
 * Reuse the last bundle while none of its input files changed, so edits still take effect without a restart. */
export async function buildDevelopmentParser(){
 if(cached&&cached.inputs.every(([file,mtime])=>{try{return statSync(file).mtimeMs===mtime;}catch{return false;}}))return cached.text;
 const result=await build({entryPoints:[fileURLToPath(new URL('./office-parser-child.ts',import.meta.url))],write:false,bundle:true,platform:'node',format:'cjs',target:'node22',metafile:true});
 const inputs=Object.keys(result.metafile.inputs).map(file=>resolve(file)).map(file=>{try{return [file,statSync(file).mtimeMs] as [string,number];}catch{return [file,-1] as [string,number];}});
 cached={inputs,text:result.outputFiles[0].text};return cached.text;
}
