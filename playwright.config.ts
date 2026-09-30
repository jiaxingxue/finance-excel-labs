// E2E tests run against the production build served under the Pages sub-path (PRD §17.4),
// so base-path bugs fail here instead of in production. Build first: `npm run build:pages`.

import { existsSync, readFileSync } from 'node:fs';
import { defineConfig, devices } from '@playwright/test';

const basePath = process.env.BASE_PATH ?? '/finance-excel-labs/';
const port = 4173;

if (!existsSync('dist/index.html')) {
  throw new Error('dist/ not found. Run `npm run build:pages` before `npm run test:e2e`.');
}

// A build made with `npm run build` (base "/") would fail every test with confusing 404s,
// so check the base that dist/ was built with: it prefixes the entry script's URL.
const builtBase = /<script[^>]*\ssrc="([^"]*)assets\//.exec(
  readFileSync('dist/index.html', 'utf8'),
)?.[1];
if (builtBase !== basePath) {
  throw new Error(
    `dist/ was built with base "${builtBase ?? 'unknown'}", but e2e tests need "${basePath}". ` +
      'Run `npm run build:pages` before `npm run test:e2e`.',
  );
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
