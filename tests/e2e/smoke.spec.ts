import { expect, test } from '@playwright/test';
import { watch } from './helpers.ts';

test('home page loads under the base path with content from the Labs document', async ({
  page,
  baseURL,
}) => {
  const seen = watch(page, baseURL!);
  await page.goto('./');
  await expect(
    page.getByRole('heading', { level: 1, name: 'Excel Labs for Financial Analysis' }),
  ).toBeVisible();
  await expect(page.getByRole('listitem')).toHaveCount(12);
  await expect(page.getByText('Budget vs. Actual Variance Report')).toBeVisible();
  await expect(page.locator('link[rel="icon"]')).toHaveAttribute('href', /^\/finance-excel-labs\//);
  expect(seen.foreign).toEqual([]);
  expect(seen.errors).toEqual([]);
});

test('a hash deep link survives a hard refresh', async ({ page, baseURL }) => {
  const seen = watch(page, baseURL!);
  await page.goto('./#/about');
  await expect(page.getByRole('heading', { level: 1, name: 'About' })).toBeVisible();
  await page.reload();
  await expect(page.getByRole('heading', { level: 1, name: 'About' })).toBeVisible();
  await expect(page).toHaveTitle('About · Excel Labs');
  expect(seen.foreign).toEqual([]);
  expect(seen.errors).toEqual([]);
});

test('navigation works and unknown routes show a not-found page', async ({ page }) => {
  await page.goto('./');
  await page.getByRole('link', { name: 'About' }).click();
  await expect(page).toHaveURL(/#\/about$/);
  await page.goto('./#/no-such-page');
  await expect(page.getByRole('heading', { level: 1, name: 'Page not found' })).toBeVisible();
});

test('404.html is the app shell, so unknown paths on GitHub Pages still boot the app', async ({
  request,
}) => {
  const [index, notFound] = await Promise.all([request.get('index.html'), request.get('404.html')]);
  expect(notFound.status()).toBe(200);
  expect(await notFound.text()).toBe(await index.text());
});
