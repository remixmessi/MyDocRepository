import {defineConfig} from '@playwright/test';
import {mkdtempSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
const scratch=mkdtempSync(join(tmpdir(),'mydoc-agent-ui-'));
export default defineConfig({testDir:'./tests/e2e',testMatch:'agent.spec.ts',workers:1,use:{channel:'chrome',baseURL:'http://127.0.0.1:4195'},webServer:{command:`MYDOC_TEST_STORAGE=1 MYDOC_DB_PATH=${scratch}/db.sqlite corepack yarn dev --port 4195`,url:'http://127.0.0.1:4195',reuseExistingServer:false}});
