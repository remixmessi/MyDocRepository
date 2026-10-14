import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests/e2e',
  testMatch: 'cloud.spec.ts',
  timeout: 180000,
  workers: 1,
  use: {
    channel: 'chromium',
    actionTimeout: 15000,
    navigationTimeout: 20000,
    trace: 'retain-on-failure',
  },
  webServer: {
    command: 'MYDOC_CONFIG_PATH=/tmp/mydoc-cloud-e2e-config.json MYDOC_CLOUD_DB_PATH=data/cloud-e2e.sqlite MYDOC_PORT=4180 node packages/frontend/apps/web/server-dist/main.cjs',
    url: 'http://127.0.0.1:4180/api/hosting',
    reuseExistingServer: false,
  },
  reporter: 'list',
});
