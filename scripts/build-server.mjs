import {build} from 'esbuild';
import {cp} from 'node:fs/promises';
const parser=await build({entryPoints:['packages/frontend/apps/web/server/office-parser-child.ts'],write:false,bundle:true,platform:'node',format:'cjs',target:'node22'});
await build({entryPoints:['packages/frontend/apps/web/server/main.ts'],outfile:'packages/frontend/apps/web/server-dist/main.cjs',bundle:true,platform:'node',format:'cjs',target:'node22',external:['better-sqlite3'],define:{__MYDOC_OFFICE_PARSER_SOURCE__:JSON.stringify(parser.outputFiles[0].text)},sourcemap:true});
// The PDF export Unicode font is read at runtime from fonts/ next to the bundle.
await cp('packages/frontend/apps/web/server/fonts','packages/frontend/apps/web/server-dist/fonts',{recursive:true});
