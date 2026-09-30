// Explore mode end to end (PRD §11 #3): every lab loads with its checks green, lesson cell links
// select cells, and the formula bar and inspector work. Expected values come from src/data/.

import { AxeBuilder } from '@axe-core/playwright';
import { expect, test, type Page } from '@playwright/test';
import assertionsJson from '../../src/data/assertions.json' with { type: 'json' };
import exercisesJson from '../../src/data/exercises.json' with { type: 'json' };
import type { Assertion, Exercise } from '../../src/content/types.ts';
import { labs } from '../../src/config/labs.config.ts';
import { watch } from './helpers.ts';

const assertions = assertionsJson as Assertion[];
const exercises = exercisesJson as Exercise[];

const nameBox = (page: Page) => page.getByLabel('Selected cell');
const formulaBar = (page: Page) => page.getByLabel('Formula bar');
const progress = (page: Page) => page.getByText(/^Checks \d+\/\d+/);
const lesson = (page: Page) => page.getByRole('region', { name: 'Lesson' });
const cell = (page: Page, sheet: string, address: string) =>
  page.locator(`[id="cell-${sheet}-${address}"]`);
const inspectorFact = (page: Page, term: string) => page.locator(`dt:text-is("${term}") + dd`);

async function openLab(page: Page, n: number) {
  await page.goto(`./#/lab/${n}`);
  await expect(
    page.getByRole('heading', { level: 1, name: new RegExp(`^Lab ${n} — `) }),
  ).toBeVisible();
}

test.describe('every lab loads in Explore mode', () => {
  for (const lab of labs) {
    test(`Lab ${lab.n}`, async ({ page, baseURL }) => {
      const seen = watch(page, baseURL!);
      await openLab(page, lab.n);
      await expect(lesson(page)).toBeVisible();
      const owned = assertions.filter((a) => lab.primarySheets.includes(a.sheet)).length;
      if (lab.sheets.length === 0) {
        await expect(page.getByRole('grid')).toHaveCount(0);
      } else if (owned === 0) {
        await expect(page.getByRole('grid')).toBeVisible();
        await expect(page.getByText(`Lab ${lab.n} owns no checks.`)).toBeVisible();
      } else {
        await expect(progress(page)).toHaveText(`Checks ${owned}/${owned} ✓`);
        const checks = page
          .getByRole('list', { name: 'Checks for this lab' })
          .getByRole('listitem');
        await expect(checks).toHaveCount(owned);
        await expect(page.getByRole('button', { name: /: fail$/ })).toHaveCount(0);
      }
      expect(seen.foreign).toEqual([]);
      expect(seen.errors).toEqual([]);
    });
  }
});

test('the lesson renders before the formula engine has loaded', async ({ page }) => {
  let release!: () => void;
  const held = new Promise<void>((resolve) => (release = resolve));
  await page.route(/engineBundle-[^/]*\.js$/, async (route) => {
    await held;
    await route.continue();
  });
  await openLab(page, 2);
  await expect(lesson(page).getByRole('heading', { name: 'Learning goals' })).toBeVisible();
  await expect(
    page.getByRole('status').filter({ hasText: 'Loading the formula engine…' }),
  ).toBeVisible();
  await expect(page.getByRole('grid')).toHaveCount(0);
  release();
  await expect(page.getByRole('grid')).toBeVisible();
  await expect(progress(page)).toContainText('✓');
});

test.describe('lesson cell links (GR-11)', () => {
  test('a qualified link selects the cell and shows its formula', async ({ page }) => {
    await openLab(page, 2);
    await expect(page.getByRole('grid')).toBeVisible();
    await lesson(page).getByRole('button', { name: 'BvA!C6' }).first().click();
    await expect(nameBox(page)).toHaveText('BvA!C6');
    const c6 = exercises.find((e) => e.sheet === 'BvA' && e.cell === 'C6')!;
    await expect(formulaBar(page)).toHaveValue(c6.formula);
    await expect(cell(page, 'BvA', 'C6')).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByText('Pattern cell')).toBeVisible();
    await expect(inspectorFact(page, 'Tier')).toHaveText('Tier A · Classic');
  });

  test('a bare link resolves to the primary sheet; ranges select every cell', async ({ page }) => {
    await openLab(page, 2);
    await expect(page.getByRole('grid')).toBeVisible();
    await lesson(page).getByRole('button', { name: 'D14', exact: true }).first().click();
    await expect(nameBox(page)).toHaveText('BvA!D14');
    await lesson(page).getByRole('button', { name: 'A5:H5', exact: true }).first().click();
    await expect(nameBox(page)).toHaveText('BvA!A5:H5');
    await expect(page.locator('[role="gridcell"][aria-selected="true"]')).toHaveCount(8);
  });

  test('a link to another sheet switches tabs; $ fragments stay plain code', async ({ page }) => {
    await openLab(page, 2);
    await expect(page.getByRole('grid')).toBeVisible();
    await lesson(page).getByRole('button', { name: 'Map!A2:A7' }).first().click();
    await expect(nameBox(page)).toHaveText('Map!A2:A7');
    await expect(page.getByRole('button', { name: 'Map', pressed: true })).toBeVisible();
    await expect(lesson(page).getByRole('button', { name: '$A6' })).toHaveCount(0);
    await expect(
      lesson(page)
        .locator('code', { hasText: /^\$A6$/ })
        .first(),
    ).toBeVisible();
  });

  test('a multi-primary lab (Lab 7) links only qualified references', async ({ page }) => {
    await openLab(page, 7);
    await expect(page.getByRole('grid')).toBeVisible();
    const links = lesson(page).locator('button[title^="Select "]');
    expect(await links.count()).toBeGreaterThan(0);
    for (const name of await links.allTextContents()) expect(name).toMatch(/^\w+!/);
  });
});

test.describe('display rule (OPEN_ISSUES #19)', () => {
  test('Checks!C8 shows 1 in the grid and 15 significant digits in the inspector', async ({
    page,
  }) => {
    await openLab(page, 12);
    await cell(page, 'Checks', 'C8').click();
    await expect(nameBox(page)).toHaveText('Checks!C8');
    await expect(cell(page, 'Checks', 'C8')).toHaveText('1');
    await expect(inspectorFact(page, 'Full value')).toHaveText('1');
  });

  test('floating-point noise from a typed formula is hidden in grid and formula bar', async ({
    page,
  }) => {
    await openLab(page, 2);
    await cell(page, 'BvA', 'J1').click();
    await page.keyboard.type('0.1');
    await page.keyboard.press('Enter');
    await cell(page, 'BvA', 'K1').click();
    await page.keyboard.type('=J1+0.2');
    await page.keyboard.press('Enter');
    await expect(cell(page, 'BvA', 'K1')).toHaveText('0.3');
    await cell(page, 'BvA', 'K1').click();
    await expect(formulaBar(page)).toHaveValue('=J1+0.2');
    await expect(inspectorFact(page, 'Full value')).toHaveText('0.3');
    await cell(page, 'BvA', 'J1').click();
    await expect(formulaBar(page)).toHaveValue('0.1');
  });
});

test('editing recalculates live, flashes changes, and undo/reset restore the checks', async ({
  page,
}) => {
  await openLab(page, 2);
  const owned = assertions.filter((a) => a.sheet === 'BvA').length;
  await expect(progress(page)).toHaveText(`Checks ${owned}/${owned} ✓`);
  await cell(page, 'BvA', 'B3').click();
  await page.keyboard.type('1');
  await page.keyboard.press('Enter');
  await expect(nameBox(page)).toHaveText('BvA!B4');
  await expect(progress(page)).not.toHaveText(`Checks ${owned}/${owned} ✓`);
  await expect(page.getByRole('button', { name: /: fail$/ }).first()).toBeVisible();
  await expect(page.locator('[class*="flash"]').first()).toBeAttached();

  await page.keyboard.press('Control+z');
  await expect(progress(page)).toHaveText(`Checks ${owned}/${owned} ✓`);

  await cell(page, 'BvA', 'B3').dblclick();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('1');
  await page.keyboard.press('Enter');
  await expect(progress(page)).not.toHaveText(`Checks ${owned}/${owned} ✓`);
  await page.getByRole('button', { name: 'Reset lab' }).click();
  await expect(progress(page)).toHaveText(`Checks ${owned}/${owned} ✓`);
});

test('an unknown sheet in a formula is rejected with a message, not a crash', async ({ page }) => {
  await openLab(page, 2);
  await cell(page, 'BvA', 'J1').click();
  // Sheet names are case-insensitive (Gl!… would resolve to GL), so use a name that doesn't exist.
  await page.keyboard.type('=SUM(Ledger!E2:E37)');
  await page.keyboard.press('Enter');
  await expect(page.getByRole('alert')).toContainText('Ledger');
  await expect(nameBox(page)).toHaveText('BvA!J1');
  await expect(page.getByRole('grid')).toBeVisible();
});

test('keyboard navigation and sheet switching (GR-13)', async ({ page }) => {
  await openLab(page, 2);
  await cell(page, 'BvA', 'A1').click();
  await page.keyboard.press('ArrowRight');
  await expect(nameBox(page)).toHaveText('BvA!B1');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox(page)).toHaveText('BvA!B3');
  await page.keyboard.press('Control+ArrowDown'); // B3 is the last input; jump to the next block
  await expect(nameBox(page)).toHaveText('BvA!B5');
  await page.keyboard.press('Home');
  await expect(nameBox(page)).toHaveText('BvA!A5');
  await page.keyboard.press('Control+PageDown');
  await expect(page.getByRole('button', { name: 'GL', pressed: true })).toBeVisible();
});

test('other labs’ sheets are read-only behind "All sheets" (GR-1)', async ({ page }) => {
  await openLab(page, 4);
  await expect(page.getByRole('grid')).toBeVisible();
  await page.getByLabel('All sheets').check();
  await page.getByRole('button', { name: /^GL/ }).click();
  await cell(page, 'GL', 'E2').click();
  await page.keyboard.type('5');
  await expect(page.getByRole('alert')).toContainText('read-only');
});

test('Lab 2 has no serious or critical accessibility violations (axe-core)', async ({ page }) => {
  await openLab(page, 2);
  await expect(page.getByRole('grid')).toBeVisible();
  await lesson(page).getByRole('button', { name: 'BvA!C6' }).first().click();
  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'])
    .analyze();
  const serious = results.violations
    .filter((v) => v.impact === 'serious' || v.impact === 'critical')
    .map((v) => `${v.id}: ${v.help} (${v.nodes.map((n) => n.target.join(' ')).join(', ')})`);
  expect(serious).toEqual([]);
});

test.describe('tablet width: the lesson is a drawer', () => {
  test.use({ viewport: { width: 1000, height: 800 } });

  test('opening the drawer and following a link selects the cell and closes it', async ({
    page,
  }) => {
    await openLab(page, 2);
    await expect(page.getByRole('grid')).toBeVisible();
    await expect(lesson(page)).toBeHidden();
    await page.getByRole('button', { name: 'Lesson', exact: true }).click();
    await expect(lesson(page)).toBeVisible();
    await lesson(page).getByRole('button', { name: 'BvA!D6' }).first().click();
    await expect(nameBox(page)).toHaveText('BvA!D6');
    await expect(lesson(page)).toBeHidden();
  });
});

test.describe('phone width', () => {
  test.use({ viewport: { width: 390, height: 844 }, hasTouch: true });

  test('one panel at a time, no sideways page scroll, links jump to the sheet', async ({
    page,
  }) => {
    await openLab(page, 2);
    const tabs = page.getByRole('group', { name: 'Workspace panels' });
    await expect(tabs).toBeVisible();
    await expect(lesson(page)).toBeVisible();
    await expect(page.getByRole('grid')).toBeHidden();
    const overflow = await page
      .locator('html')
      .evaluate((html: HTMLElement) => html.scrollWidth - html.clientWidth);
    expect(overflow).toBeLessThanOrEqual(0);

    await lesson(page).getByRole('button', { name: 'BvA!C6' }).first().click();
    await expect(tabs.getByRole('button', { name: 'Sheet', pressed: true })).toBeVisible();
    await expect(page.getByRole('grid')).toBeVisible();
    await expect(nameBox(page)).toHaveText('BvA!C6');

    await tabs.getByRole('button', { name: 'Checks' }).click();
    await expect(page.getByRole('list', { name: 'Checks for this lab' })).toBeVisible();
    await expect(page.getByRole('grid')).toBeHidden();
  });
});
