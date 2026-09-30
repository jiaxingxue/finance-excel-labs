// What-if mode end to end (PRD §6.4, §11 #5): every Appendix B.4 control exists and moves its cell,
// and every Appendix E experiment reproduces its documented values when driven through the
// controls. Expected values come from src/data/experiments.json; controls from labs.config.ts.

import { AxeBuilder } from '@axe-core/playwright';
import { expect, test, type Locator, type Page } from '@playwright/test';
import experimentsJson from '../../src/data/experiments.json' with { type: 'json' };
import type { Experiment } from '../../src/content/types.ts';
import { labs, type Control } from '../../src/config/labs.config.ts';
import { parseRef } from '../../src/engine/address.ts';
import { controlCells } from '../../src/whatif/controls.ts';
import { watch } from './helpers.ts';

const experiments = experimentsJson as unknown as Experiment[];

const whatIf = (page: Page) => page.getByRole('region', { name: 'What-if' });
const progress = (page: Page) => page.getByText(/^Checks \d+\/\d+/);
const nameBox = (page: Page) => page.getByLabel('Selected cell');
const inspectorFact = (page: Page, term: string) => page.locator(`dt:text-is("${term}") + dd`);
const gridCell = (page: Page, ref: string) => {
  const { sheet, cell } = parseRef(ref);
  return page.locator(`[id="cell-${sheet}-${cell}"]`);
};

async function openWhatIf(page: Page, n: number) {
  await page.goto(`./#/lab/${n}?mode=whatif`);
  await expect(progress(page).or(page.getByText(`Lab ${n} owns no checks.`))).toBeVisible();
  await expect(whatIf(page)).toBeVisible();
}

/** The control that writes `ref`, and the lab it belongs to. */
function controlFor(lab: number, ref: string): Control {
  const config = labs.find((l) => l.n === lab)!;
  for (const group of config.controls) {
    for (const control of group.controls) {
      if (controlCells(control, experiments).some((c) => `${c.sheet}!${c.cell}` === ref)) {
        return control;
      }
    }
  }
  throw new Error(`No control writes ${ref}`);
}

/** The typed input for a cell inside a table control (labelled with its address). */
const tableInput = (page: Page, ref: string): Locator =>
  whatIf(page).getByLabel(`(${ref})`, { exact: false });

/** Sets `ref` to `value` through its control, as a learner would. */
async function drive(page: Page, lab: number, ref: string, value: number | string) {
  const control = controlFor(lab, ref);
  const panel = whatIf(page);
  switch (control.type) {
    case 'slider': {
      const box = panel.getByRole('textbox', { name: / value$/ }).nth(sliderIndex(lab, ref));
      await box.fill(String(value));
      await box.press('Enter');
      break;
    }
    case 'number':
      await panel.locator('input[type="number"]').nth(numberIndex(lab, ref)).fill(String(value));
      break;
    case 'select':
      await panel.getByRole('combobox').nth(selectIndex(lab, ref)).selectOption(String(value));
      break;
    case 'segmented':
      await panel
        .getByRole('radio', { name: control.options.find((o) => o.value === value)!.label })
        .check();
      break;
    case 'date':
      await panel.locator('input[type="date"]').fill(String(value));
      break;
    case 'table': {
      const details = panel.locator('details');
      if ((await details.count()) > 0 && !(await details.first().getAttribute('open'))) {
        await details.first().locator('summary').click();
      }
      await tableInput(page, ref).fill(String(value));
      break;
    }
    case 'experiment':
      await panel.getByRole('button', { name: control.label }).click();
      break;
  }
}

/** Position of a control among the lab's controls of the same kind, in panel order. */
function indexAmong(lab: number, ref: string, type: Control['type']): number {
  const config = labs.find((l) => l.n === lab)!;
  const cells = config.controls
    .flatMap((g) => g.controls)
    .filter((c) => c.type === type)
    .map((c) => ('cell' in c ? c.cell : ''));
  const i = cells.indexOf(ref);
  if (i < 0) throw new Error(`${ref} is not a ${type} control`);
  return i;
}
const sliderIndex = (lab: number, ref: string) => indexAmong(lab, ref, 'slider');
const numberIndex = (lab: number, ref: string) => indexAmong(lab, ref, 'number');
const selectIndex = (lab: number, ref: string) => indexAmong(lab, ref, 'select');

/** Reads a cell through the UI: select it from the key outputs, then read the inspector. */
async function readCell(page: Page, ref: string): Promise<string | number | null> {
  await whatIf(page)
    .getByRole('button', { name: new RegExp(`^${ref.replace('!', '!')}\\b`) })
    .click();
  await expect(nameBox(page)).toHaveText(ref);
  const full = inspectorFact(page, 'Full value');
  if ((await full.count()) > 0) return Number(await full.textContent());
  const value = inspectorFact(page, 'Value');
  if ((await value.count()) === 0) return null;
  return (await value.textContent()) ?? '';
}

function matches(actual: string | number | null, expected: number | string | boolean | null) {
  if (expected === null || expected === '') return actual === null || actual === '';
  if (typeof expected === 'boolean') return actual === (expected ? 'TRUE' : 'FALSE');
  if (typeof expected === 'number') {
    // The inspector shows 15 significant digits; experiments allow 1e-6 relative or 0.005.
    return (
      typeof actual === 'number' &&
      Math.abs(actual - expected) <= Math.max(Math.abs(expected) * 1e-6, 0.005)
    );
  }
  return actual === expected;
}

test.describe('mode tabs', () => {
  test('?mode=whatif is a deep link, and Explore hides the controls', async ({ page }) => {
    await openWhatIf(page, 2);
    await expect(page.getByRole('button', { name: 'What-if', exact: true })).toHaveAttribute(
      'aria-pressed',
      'true',
    );
    await page.getByRole('button', { name: 'Explore', exact: true }).click();
    await expect(whatIf(page)).toHaveCount(0);
    await expect(page).not.toHaveURL(/mode=/);
    await page.getByRole('button', { name: 'What-if', exact: true }).click();
    await expect(page).toHaveURL(/mode=whatif/);
    await expect(whatIf(page)).toBeVisible();
  });

  for (const lab of labs.filter((l) => l.controls.length === 0)) {
    test(`Lab ${lab.n} has no controls, so What-if is disabled`, async ({ page }) => {
      await page.goto(`./#/lab/${lab.n}?mode=whatif`);
      await expect(page.getByRole('button', { name: /^What-if/ })).toBeDisabled();
      await expect(whatIf(page)).toHaveCount(0);
    });
  }
});

test.describe('every control moves its cell (WI-1, WI-2)', () => {
  for (const lab of labs.filter((l) => l.controls.length > 0)) {
    test(`Lab ${lab.n}`, async ({ page, baseURL }) => {
      const seen = watch(page, baseURL!);
      await openWhatIf(page, lab.n);
      for (const control of lab.controls.flatMap((g) => g.controls)) {
        const refs = controlCells(control, experiments).map((c) => `${c.sheet}!${c.cell}`);
        const ref = refs[0]!;
        // Each control's cells are on the lab's first sheet, which is on screen.
        const before = await gridCell(page, ref).textContent();
        const next = nextValue(control, ref, before ?? '');
        await drive(page, lab.n, ref, next);
        await expect(gridCell(page, ref), `${ref} via ${control.type}`).not.toHaveText(
          before ?? '',
        );
        await whatIf(page).getByRole('button', { name: 'Reset inputs' }).click();
        await expect(gridCell(page, ref)).toHaveText(before ?? '');
      }
      await expect(progress(page).or(page.getByText(/owns no checks/))).not.toContainText('✗');
      expect(seen.errors).toEqual([]);
    });
  }
});

/** A value different from the current one, inside the control's range. */
function nextValue(control: Control, ref: string, shown: string): number | string {
  const n = Number(shown.replace(/[,%()]/g, ''));
  switch (control.type) {
    case 'slider':
      return String(control.min === 0 ? control.max : control.min);
    case 'number':
      return String((Number.isFinite(n) ? n : 0) + control.step * 2);
    case 'select':
      return control.options.find((o) => o !== shown)!;
    case 'segmented':
      return control.options.find((o) => String(o.value) !== shown)!.value;
    case 'date':
      return control.max;
    case 'table':
      return String((Number.isFinite(n) ? n : 0) + 10);
    case 'experiment':
      return ref;
  }
}

test.describe('experiments reproduce their documented values through the controls (§11 #5)', () => {
  for (const e of experiments) {
    test(`${e.id} ${e.title}`, async ({ page }) => {
      await openWhatIf(page, e.lab);
      const control = Object.keys(e.changes).map((ref) => controlFor(e.lab, ref));
      if (control[0]!.type === 'experiment') {
        await drive(page, e.lab, Object.keys(e.changes)[0]!, '');
      } else {
        for (const [ref, value] of Object.entries(e.changes)) await drive(page, e.lab, ref, value);
      }
      for (const [ref, expected] of Object.entries(e.expect)) {
        let last: string | number | null = null;
        await expect
          .poll(
            async () => {
              last = await readCell(page, ref);
              return matches(last, expected);
            },
            { message: `${e.id}: ${ref} should be ${JSON.stringify(expected)}` },
          )
          .toBe(true)
          .catch((error: unknown) => {
            throw new Error(`${String(error)}
Last value read: ${JSON.stringify(last)}`);
          });
      }
    });
  }
});

test('Lab 9: "Inject outlier" is undone by a second press (WI-3)', async ({ page }) => {
  await openWhatIf(page, 9);
  const e91 = experiments.find((e) => e.id === 'E9.1')!;
  const [ref] = Object.keys(e91.changes) as [string];
  const before = await gridCell(page, ref).textContent();
  await whatIf(page).getByRole('button', { name: 'Inject outlier' }).click();
  await expect(whatIf(page).getByRole('button', { name: 'Remove outlier' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await expect(gridCell(page, ref)).not.toHaveText(before ?? '');
  await whatIf(page).getByRole('button', { name: 'Remove outlier' }).click();
  await expect(gridCell(page, ref)).toHaveText(before ?? '');
  await expect(progress(page)).not.toContainText('✗');
});

test('key outputs show before → now, and a note explains failing checks', async ({ page }) => {
  await openWhatIf(page, 2);
  const e21 = experiments.find((e) => e.id === 'E2.1')!;
  const [ref, value] = Object.entries(e21.changes)[0]!;
  await expect(whatIf(page).getByText(/Checks compare/)).toHaveCount(0);
  await drive(page, 2, ref, value);
  await expect(whatIf(page).getByText(/Checks compare/)).toBeVisible();
  const outputs = whatIf(page).getByRole('list', { name: 'Key outputs, before and now' });
  await expect(outputs.getByRole('listitem')).toHaveCount(
    new Set(experiments.filter((x) => x.lab === 2).flatMap((x) => Object.keys(x.expect))).size,
  );
  await expect(outputs.getByText('→').first()).toBeVisible();
});

test('conditional formats follow the data (GR-12)', async ({ page }) => {
  await openWhatIf(page, 7);
  const e72 = experiments.find((e) => e.id === 'E7.2')!;
  const status = gridCell(page, 'BankRec!B20');
  const before = await status.evaluate((el) => getComputedStyle(el).backgroundColor);
  for (const [ref, value] of Object.entries(e72.changes)) await drive(page, 7, ref, value);
  await expect
    .poll(() => status.evaluate((el) => getComputedStyle(el).backgroundColor))
    .not.toBe(before);
});

test('axe: no serious or critical violations in What-if mode', async ({ page }) => {
  await openWhatIf(page, 10);
  const results = await new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'])
    .analyze();
  const serious = results.violations.filter((v) => ['serious', 'critical'].includes(v.impact!));
  expect(serious.map((v) => `${v.id}: ${v.nodes.map((n) => n.target).join(', ')}`)).toEqual([]);
});
