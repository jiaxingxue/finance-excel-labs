// What-if mode (WI-1–WI-4): controls bound to the lab's input cells, "Reset inputs", and the key
// output cells with their before → now values. Every control writes through the store, so the
// grid recalculates and flashes changed cells (GR-6).

import { useMemo } from 'react';
import experimentsJson from '../../data/experiments.json' with { type: 'json' };
import type { Experiment } from '../../content/types.ts';
import type { LabConfig } from '../../config/labs.config.ts';
import type { CellEntry, CellRef, CellValue } from '../../engine/types.ts';
import { formatCell } from '../../format/displayValue.ts';
import { groupTitleCell, keyOutputs, labControlCells, sameValue } from '../../whatif/controls.ts';
import type { WorkbookStore } from '../../workspace/workbookStore.ts';
import side from '../side/Side.module.css';
import { ControlView } from './controls.tsx';
import styles from './WhatIf.module.css';

const experiments = experimentsJson as unknown as Experiment[];

interface WhatIfPanelProps {
  lab: LabConfig;
  store: WorkbookStore;
  /** Store version; a new value re-renders after recalculation. */
  version: number;
  /** Writes cells; returns an error message, or null when accepted. */
  onWrite: (entries: CellEntry[]) => string | null;
  onSelectCell: (ref: CellRef) => void;
}

export function WhatIfPanel({ lab, store, version, onWrite, onSelectCell }: WhatIfPanelProps) {
  const inputs = useMemo(() => labControlCells(lab, experiments), [lab]);
  const outputs = useMemo(() => keyOutputs(lab.n, experiments), [lab]);
  const text = (ref: CellRef) => formatCell(store.getValue(ref));
  const changedInputs = inputs.filter(
    (ref) => !sameValue(store.getValue(ref), store.baseValue(ref)),
  );

  return (
    <section className={`${side.panel} ${styles.panel}`} aria-labelledby="whatif-heading">
      <h2 id="whatif-heading" className={side.heading}>
        What-if
        <button
          type="button"
          className={styles.reset}
          disabled={changedInputs.length === 0}
          onClick={() =>
            onWrite(changedInputs.map((ref) => ({ ref, content: store.originalContent(ref) })))
          }
        >
          Reset inputs
        </button>
      </h2>

      {changedInputs.length > 0 && (
        <p className={styles.note} role="status">
          Checks compare with the Labs document’s base inputs, so some may fail while an input is
          changed. Reset inputs to restore them.
        </p>
      )}

      {lab.controls.map((group, i) => {
        const title = groupTitleCell(group);
        return (
          <fieldset key={i} className={styles.group}>
            {title ? (
              <legend className={styles.legend}>{text(title)}</legend>
            ) : (
              <legend className={styles.srOnly}>Inputs</legend>
            )}
            {group.controls.map((control, j) => (
              <ControlView
                key={j}
                control={control}
                store={store}
                version={version}
                experiments={experiments}
                onWrite={onWrite}
                context={title ? text(title) : undefined}
              />
            ))}
          </fieldset>
        );
      })}

      <h3 className={styles.subheading}>Key outputs</h3>
      {outputs.length === 0 ? (
        <p className={side.muted}>
          Lab {lab.n} has no experiments, so no key outputs are listed. Watch the grid.
        </p>
      ) : (
        <ul className={styles.deltas} aria-label="Key outputs, before and now">
          {outputs.map((ref) => {
            const fmt = store.fmt(ref);
            const before = store.baseValue(ref);
            const now = store.getValue(ref);
            const changed = !sameValue(before, now);
            const show = (v: CellValue) => formatCell(v, fmt) || '(blank)';
            return (
              <li key={`${ref.sheet}!${ref.cell}`}>
                <button
                  type="button"
                  className={changed ? styles.changed : undefined}
                  onClick={() => onSelectCell(ref)}
                >
                  <code>
                    {ref.sheet}!{ref.cell}
                  </code>
                  <span className={styles.values}>
                    {changed ? (
                      <>
                        {show(before)} <span aria-label="changed to">→</span>{' '}
                        <strong>{show(now)}</strong>
                      </>
                    ) : (
                      <>
                        {show(now)} <span className={styles.unchanged}>(unchanged)</span>
                      </>
                    )}
                  </span>
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
}
