// Formula bar commit on blur, as in Excel, and the hover tooltip on cut-off labels (M3a).

import { expect, test, type Page } from '@playwright/test';
import workbookJson from '../../src/data/workbook.json' with { type: 'json' };
import type { WorkbookSpec } from '../../src/content/types.ts';

const workbook = workbookJson as WorkbookSpec;
const formulaBar = (page: Page) => page.getByLabel('Formula bar');
const nameBox = (page: Page) => page.getByLabel('Selected cell');
const cell = (page: Page, sheet: string, address: string) =>
  page.locator(`[id="cell-${sheet}-${address}"]`);

async function openLab2(page: Page) {
  await page.goto('./#/lab/2');
  await expect(page.getByRole('grid')).toBeVisible();
}

test('an unsaved formula-bar edit commits to its own cell when you click another cell', async ({
  page,
}) => {
  await openLab2(page);
  await cell(page, 'BvA', 'J2').click();
  await formulaBar(page).fill('=1+2');
  await cell(page, 'BvA', 'J4').click();
  await expect(nameBox(page)).toHaveText('BvA!J4');
  await expect(cell(page, 'BvA', 'J2')).toHaveText('3');
  await expect(cell(page, 'BvA', 'J4')).toHaveText('');
});

test('an unsaved edit also commits when you switch sheets', async ({ page }) => {
  await openLab2(page);
  await cell(page, 'BvA', 'J2').click();
  await formulaBar(page).fill('kept');
  await page.getByRole('button', { name: 'GL', exact: true }).click();
  await page.getByRole('button', { name: 'BvA', exact: true }).click();
  await expect(cell(page, 'BvA', 'J2')).toHaveText('kept');
});

test('Esc discards the edit, and nothing commits on the blur that follows', async ({ page }) => {
  await openLab2(page);
  await cell(page, 'BvA', 'J2').click();
  await formulaBar(page).fill('discard me');
  await formulaBar(page).press('Escape');
  await cell(page, 'BvA', 'J4').click();
  await expect(cell(page, 'BvA', 'J2')).toHaveText('');
});

test('a rejected entry on blur explains why', async ({ page }) => {
  await openLab2(page);
  await cell(page, 'BvA', 'J2').click();
  await formulaBar(page).fill('=SUM(Ledger!E2:E37)');
  await cell(page, 'BvA', 'J4').click();
  await expect(page.getByRole('alert')).toContainText('unknown sheet "Ledger"');
  await expect(cell(page, 'BvA', 'J2')).toHaveText('');
});

test('a cut-off label shows its full text on hover; a label that fits does not', async ({
  page,
}) => {
  await openLab2(page);
  // Pick the longest and shortest labels in column A from the Labs document.
  const labels = Object.entries(workbook.sheets.BvA!)
    .filter(([address, spec]) => /^A\d+$/.test(address) && 'v' in spec)
    .map(([address, spec]) => ({ address, text: String((spec as { v: unknown }).v) }))
    .sort((a, b) => a.text.length - b.text.length);
  const [short, long] = [labels[0]!, labels.at(-1)!];

  await cell(page, 'BvA', long.address).hover();
  await expect(cell(page, 'BvA', long.address).locator('span')).toHaveAttribute('title', long.text);
  await cell(page, 'BvA', short.address).hover();
  await expect(cell(page, 'BvA', short.address).locator('span')).not.toHaveAttribute('title');
});
