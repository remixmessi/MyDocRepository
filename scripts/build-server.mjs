import {build} from 'esbuild';
await build({entryPoints:['packages/frontend/apps/web/server/main.ts'],outfile:'packages/frontend/apps/web/server-dist/main.cjs',bundle:true,platform:'node',format:'cjs',target:'node22',external:['better-sqlite3'],sourcemap:true});
