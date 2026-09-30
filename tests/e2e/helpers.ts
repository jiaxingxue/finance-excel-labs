import type { Page } from '@playwright/test';

/** Collects console errors and any request that leaves the site's own origin (PRD §8 privacy). */
export function watch(page: Page, baseURL: string) {
  const origin = new URL(baseURL).origin;
  const foreign: string[] = [];
  const errors: string[] = [];
  page.on('request', (req) => {
    const url = req.url();
    if (!url.startsWith('data:') && new URL(url).origin !== origin) foreign.push(url);
  });
  page.on('console', (msg) => {
    if (msg.type() === 'error') errors.push(msg.text());
  });
  page.on('pageerror', (err) => errors.push(err.message));
  return { foreign, errors };
}
