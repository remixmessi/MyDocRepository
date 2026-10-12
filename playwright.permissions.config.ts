import {defineConfig} from '@playwright/test';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
const database=join(tmpdir(),'mydoc-permissions-'+Date.now()+'.sqlite');
export default defineConfig({testDir:'./tests/e2e',testMatch:'document-permissions.spec.ts',timeout:90000,workers:1,use:{channel:process.env.MYDOC_TEST_BROWSER??'chrome'},webServer:{command:`MYDOC_CLOUD_DB_PATH=${database} MYDOC_PORT=4181 node packages/frontend/apps/web/server-dist/main.cjs`,url:'http://127.0.0.1:4181/api/hosting',reuseExistingServer:false},reporter:'list'});
