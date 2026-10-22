import {defineConfig} from '@playwright/test';
export default defineConfig({testDir:'./tests/e2e',testMatch:['office-text.spec.ts','office-shell.spec.ts'],timeout:60000,workers:1,use:{channel:'chromium',baseURL:process.env.MYDOC_OFFICE_TEST_URL??'http://127.0.0.1:4185',trace:'retain-on-failure'},reporter:'list'});
