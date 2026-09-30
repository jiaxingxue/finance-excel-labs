// E2E tests run against the production build served under the Pages sub-path (PRD §17.4),
// so base-path bugs fail here instead of in production. Build first: `npm run build:pages`.

import { existsSync } from 'node:fs';
import { defineConfig, devices } from '@playwright/test';

const basePath = process.env.BASE_PATH ?? '/finance-excel-labs/';
const port = 4173;

if (!existsSync('dist/index.html')) {
  throw new Error('dist/ not found. Run `npm run build:pages` before `npm run test:e2e`.');
}

export default defineConfig({
  testDir: 'tests/e2e',
  outputDir: 'test-results',
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  reporter: [['list'], ['html', { outputFolder: 'reports/playwright', open: 'never' }]],
  use: {
    baseURL: `http://localhost:${port}${basePath}`,
    trace: 'retain-on-failure',
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: {
    command: `npx vite preview --port ${port} --strictPort`,
    url: `http://localhost:${port}${basePath}`,
    env: { BASE_PATH: basePath },
    reuseExistingServer: !process.env.CI,
  },
});
