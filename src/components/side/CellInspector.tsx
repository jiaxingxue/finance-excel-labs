// EX-2: what the selected cell holds: formula, value (formatted and to 15 significant digits),
// tier badge, and whether the lesson's formula table uses it as a pattern cell.

import type { Exercise } from '../../content/types.ts';
import { formulaTier, type Tier } from '../../config/tiers.ts';
import type { CellValue } from '../../engine/types.ts';
import { formatCell, toSignificant15 } from '../../format/displayValue.ts';
import styles from './Side.module.css';

const TIER_NAMES: Record<Tier, string> = {
  A: 'Classic',
  B: 'Excel 2016/2019',
  C: 'Dynamic arrays',
  D: 'Microsoft 365 only',
};

interface CellInspectorProps {
  label: string;
  /** Null while the engine loads. */
  cell: { value: CellValue; formula: string | null; fmt?: string } | null;
  exercise: Exercise | undefined;
}

export function CellInspector({ label, cell, exercise }: CellInspectorProps) {
  const tier = cell?.formula ? formulaTier(cell.formula) : null;
  return (
    <section className={styles.panel} aria-labelledby="inspector-heading">
      <h2 id="inspector-heading" className={styles.heading}>
        Cell inspector
      </h2>
      {cell === null ? (
        <p className={styles.muted}>Loading the formula engine…</p>
      ) : (
        <dl className={styles.facts}>
          <dt>Cell</dt>
          <dd>
            <code>{label}</code>
          </dd>
          <dt>Contents</dt>
          <dd>
            {cell.formula ? (
              <code className={styles.formula}>{cell.formula}</code>
            ) : cell.value === null ? (
              'Blank'
            ) : (
              'Constant (input)'
            )}
          </dd>
          {cell.value !== null && (
            <>
              <dt>Value</dt>
              <dd>{formatCell(cell.value, cell.fmt)}</dd>
            </>
          )}
          {typeof cell.value === 'number' && (
            <>
              <dt>Full value</dt>
              <dd>
                <code>{toSignificant15(cell.value)}</code>
              </dd>
            </>
          )}
          {cell.formula && (
            <>
              <dt>Tier</dt>
              <dd>
                {tier ? (
                  <span className={styles.badge}>
                    Tier {tier} · {TIER_NAMES[tier]}
                  </span>
                ) : (
                  <span className={styles.muted}>Unknown function</span>
                )}
              </dd>
            </>
          )}
          {exercise && (
            <>
              <dt>Lesson</dt>
              <dd>
                <span className={styles.badge}>Pattern cell</span> Lab {exercise.lab} result:{' '}
                {exercise.result}
              </dd>
            </>
          )}
        </dl>
      )}
    </section>
  );
}
