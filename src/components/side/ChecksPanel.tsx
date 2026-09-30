// EX-4: the lab's assertions, graded live against the engine. Expected values come only from
// src/data/assertions.json. Pass/fail is shown with a word and an icon, never color alone.

import type { Assertion } from '../../content/types.ts';
import type { CellValue } from '../../engine/types.ts';
import { formatCell } from '../../format/displayValue.ts';
import styles from './Side.module.css';

export interface CheckRow {
  assertion: Assertion;
  actual: CellValue;
  pass: boolean;
  fmt?: string;
}

interface ChecksPanelProps {
  /** Null while the engine loads. */
  rows: CheckRow[] | null;
  /** Shown when the lab owns no assertions. */
  emptyNote: string;
  onSelect: (sheet: string, cell: string) => void;
}

export function ChecksPanel({ rows, emptyNote, onSelect }: ChecksPanelProps) {
  const passed = rows?.filter((r) => r.pass).length ?? 0;
  return (
    <section className={styles.panel} aria-labelledby="checks-heading">
      <h2 id="checks-heading" className={styles.heading}>
        Checks
        {rows && rows.length > 0 && (
          <span className={styles.count}>
            {passed} of {rows.length} pass
          </span>
        )}
      </h2>
      {rows === null ? (
        <p className={styles.muted}>Loading the formula engine…</p>
      ) : rows.length === 0 ? (
        <p className={styles.muted}>{emptyNote}</p>
      ) : (
        <ul className={styles.checks} aria-label="Checks for this lab">
          {rows.map(({ assertion, actual, pass, fmt }) => {
            const ref = `${assertion.sheet}!${assertion.cell}`;
            return (
              <li key={ref}>
                <button
                  type="button"
                  className={pass ? styles.pass : styles.fail}
                  onClick={() => onSelect(assertion.sheet, assertion.cell)}
                  aria-label={`${ref}: ${pass ? 'pass' : 'fail'}`}
                >
                  <span className={styles.status} aria-hidden="true">
                    {pass ? '✓' : '✗'}
                  </span>
                  <code className={styles.ref}>{ref}</code>
                  <span className={styles.actual}>{formatCell(actual, fmt) || '(blank)'}</span>
                  {!pass && (
                    <span className={styles.expected}>
                      expected {formatCell(assertion.expected, fmt) || '(blank)'}
                    </span>
                  )}
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
}
