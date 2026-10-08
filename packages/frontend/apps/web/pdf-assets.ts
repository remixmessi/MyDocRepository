import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { readdirSync, readFileSync } from 'node:fs';
import type { Plugin } from 'vite';

// Preserve PDF.js resource names: its worker requests CMaps/fonts by name.
export function pdfAssets(): Plugin {
 const root=dirname(createRequire(import.meta.url).resolve('pdfjs-dist/package.json'));
 const files=new Map<string,Buffer>();
 for(const folder of ['cmaps','standard_fonts','wasm','iccs']) {
  for(const name of readdirSync(join(root,folder))) {
   if(name.endsWith('.md'))continue;
   files.set(`pdf-assets/${folder}/${name}`,readFileSync(join(root,folder,name)));
  }
 }
 return {name:'local-pdf-assets',configureServer(server){server.middlewares.use((req,res,next)=>{
  const key=req.url?.split('?')[0].replace(/^\//,'');const bytes=key?files.get(key):undefined;
  if(!bytes)return next();res.setHeader('Content-Type',key!.endsWith('.wasm')?'application/wasm':'application/octet-stream');res.end(bytes);
 });},generateBundle(){for(const [fileName,source] of files)this.emitFile({type:'asset',fileName,source});}};
}
