// Milestone screenshots of the lab workspace, saved to docs/milestones/m2/. Skipped in normal
// e2e runs (they would rewrite committed images); run with `npm run screenshots:m2`.

import { expect, test, type Page } from '@playwright/test';

test.skip(!process.env.SCREENSHOTS, 'Run with `npm run screenshots:m2`.');
test.use({ colorScheme: 'light' });

const out = (name: string) => `docs/milestones/m2/${name}.png`;

/** The checks count appears once the engine has calculated (the grid may be on a hidden tab). */
const engineReady = (page: Page) => expect(page.getByText(/^Checks \d+\/\d+/)).toBeVisible();

async function openLab2AtC6(page: Page) {
  await page.goto('./#/lab/2');
  await engineReady(page);
  await page
    .getByRole('region', { name: 'Lesson' })
    .getByRole('button', { name: 'BvA!C6' })
    .first()
    .click();
  await expect(page.getByLabel('Selected cell')).toHaveText('BvA!C6');
}

test.describe('desktop', () => {
  test.use({ viewport: { width: 1440, height: 900 } });

  test('lab workspace', async ({ page }) => {
    await openLab2AtC6(page);
    await page
      .getByRole('region', { name: 'Lesson' })
      .evaluate((el: HTMLElement) => (el.scrollTop = 0));
    await page.screenshot({ path: out('desktop-lab2') });
  });
});

test.describe('phone', () => {
  test.use({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, hasTouch: true });

  test('lesson tab', async ({ page }) => {
    await page.goto('./#/lab/2');
    await engineReady(page);
    await page.screenshot({ path: out('phone-lab2-lesson') });
  });

  test('sheet tab', async ({ page }) => {
    await openLab2AtC6(page);
    await page.locator('html').evaluate((html: HTMLElement) => (html.scrollTop = 0));
    await page.screenshot({ path: out('phone-lab2-sheet') });
  });
});
