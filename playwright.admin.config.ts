import {defineConfig} from '@playwright/test';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
const database=join(tmpdir(),'mydoc-admin-'+Date.now()+'.sqlite');
export default defineConfig({testDir:'./tests/e2e',testMatch:'global-workspace-admin.spec.ts',timeout:90000,workers:1,use:{channel:process.env.MYDOC_TEST_BROWSER??'chrome',trace:'retain-on-failure'},webServer:{command:`MYDOC_CLOUD_DB_PATH=${database} MYDOC_PORT=4182 MYDOC_ADMIN_EMAIL=admin@example.com MYDOC_ADMIN_PASSWORD=admin-test-password-123 node packages/frontend/apps/web/server-dist/main.cjs`,url:'http://127.0.0.1:4182/api/hosting',reuseExistingServer:false},reporter:'list'});
