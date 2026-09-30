// public/verification.json ships with the site (PRD §17.5). GitHub Pages answers a missing file
// with 404.html, which is the app shell, so a plain "status 200" check would pass without the file.
// This test parses the body as JSON and checks the counts against src/data/.

import { expect, test } from '@playwright/test';
import manifest from '../../src/data/manifest.json' with { type: 'json' };

interface Verification {
  assertions: { passed: number; total: number };
  experiments: { passed: number; total: number };
  label: string;
}

test('verification.json is real JSON reporting every assertion and experiment as passing', async ({
  request,
}) => {
  const response = await request.get('verification.json');
  expect(response.status()).toBe(200);
  const body = await response.text();
  expect(
    body.trimStart().startsWith('{'),
    'verification.json is missing from dist/ (the server returned the HTML fallback). ' +
      'Run `npm test -- --reporter=default --reporter=json --outputFile=reports/vitest.json`, ' +
      'then `npm run verification:write`, then `npm run build:pages`.',
  ).toBe(true);
  const v = JSON.parse(body) as Verification;
  const { assertions, experiments } = manifest.counts;
  expect(v.assertions).toEqual({ passed: assertions, total: assertions });
  expect(v.experiments).toEqual({ passed: experiments, total: experiments });
  expect(v.label).toBe(`${assertions}/${assertions} passing`);
  // The counts themselves are pinned by the golden cross-check (tests/golden.test.ts).
  expect([assertions, experiments]).toEqual([277, 12]);
});
