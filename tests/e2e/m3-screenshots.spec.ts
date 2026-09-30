// Milestone 3 screenshots, saved to docs/milestones/m3/. Skipped in normal e2e runs (they would
// rewrite committed images); run with `npm run screenshots:m3`.

import { expect, test, type Page } from '@playwright/test';

test.skip(!process.env.SCREENSHOTS, 'Run with `npm run screenshots:m3`.');
test.use({ colorScheme: 'light' });

const out = (name: string) => `docs/milestones/m3/${name}.png`;
const engineReady = (page: Page) => expect(page.getByText(/^Checks \d+\/\d+/)).toBeVisible();
const whatIf = (page: Page) => page.getByRole('region', { name: 'What-if' });

test.describe('desktop', () => {
  test.use({ viewport: { width: 1440, height: 900 } });

  test('Lab 2 What-if: % threshold lowered to 3%', async ({ page }) => {
    await page.goto('./#/lab/2?mode=whatif');
    await engineReady(page);
    const box = whatIf(page).getByLabel('Pct threshold value');
    await box.fill('3%');
    await box.press('Enter');
    // Bring the before → now list into view, with the first changed output on screen.
    await whatIf(page)
      .getByRole('button', { name: /changed to/ })
      .first()
      .scrollIntoViewIfNeeded();
    await page.screenshot({ path: out('desktop-lab2-whatif') });
  });

  test('Lab 7 What-if: conditional formats on Book', async ({ page }) => {
    await page.goto('./#/lab/7?mode=whatif');
    await engineReady(page);
    await page.getByRole('button', { name: 'Book', exact: true }).click();
    await page.screenshot({ path: out('desktop-lab7-book-formats') });
  });

  test('Lab 10 What-if: Upside scenario', async ({ page }) => {
    await page.goto('./#/lab/10?mode=whatif');
    await engineReady(page);
    await whatIf(page).getByRole('radio', { name: 'Up' }).check();
    await page.screenshot({ path: out('desktop-lab10-whatif') });
  });
});

test.describe('phone', () => {
  test.use({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true });

  test('Lab 8 What-if tab', async ({ page }) => {
    await page.goto('./#/lab/8?mode=whatif');
    await engineReady(page);
    await page.getByRole('button', { name: 'What-if', exact: true }).last().click();
    await page.screenshot({ path: out('phone-lab8-whatif') });
  });
});
