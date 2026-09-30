// One What-if control per config entry (WI-1). The engine value is the source of truth: a control
// shows the cell's current value, so edits made in the grid show up here too.

import { useEffect, useId, useRef, useState } from 'react';
import type { Experiment } from '../../content/types.ts';
import type {
  Control,
  DateControl,
  ExperimentAction,
  NumberControl,
  SegmentedControl,
  SelectControl,
  SliderControl,
  TableControl,
} from '../../config/labs.config.ts';
import { isoDateToSerial, parseRef, serialToIsoDate } from '../../engine/address.ts';
import type { CellEntry, CellRef } from '../../engine/types.ts';
import { formatCell, toSignificant15 } from '../../format/displayValue.ts';
import { experimentChangeContent } from '../../grading/match.ts';
import {
  controlLabelCell,
  findExperiment,
  tableCells,
  tableLabels,
} from '../../whatif/controls.ts';
import { parseInput } from '../../workspace/parseInput.ts';
import type { WorkbookStore } from '../../workspace/workbookStore.ts';
import styles from './WhatIf.module.css';

interface ControlProps<C> {
  control: C;
  store: WorkbookStore;
  version: number;
  experiments: readonly Experiment[];
  onWrite: (entries: CellEntry[]) => string | null;
  /** The group's title, e.g. "Monthly churn", so repeated labels ("Base") stay distinct. */
  context?: string;
}

export function ControlView(props: ControlProps<Control>) {
  const { control } = props;
  switch (control.type) {
    case 'slider':
      return <Slider {...props} control={control} />;
    case 'number':
      return <NumberRow {...props} control={control} />;
    case 'select':
      return <Select {...props} control={control} />;
    case 'segmented':
      return <Segmented {...props} control={control} />;
    case 'date':
      return <DateRow {...props} control={control} />;
    case 'table':
      return <Table {...props} control={control} />;
    case 'experiment':
      return <ExperimentButton {...props} control={control} />;
  }
}

const labelText = (store: WorkbookStore, ref: CellRef) => formatCell(store.getValue(ref));
const numberAt = (store: WorkbookStore, ref: CellRef) => {
  const v = store.getValue(ref);
  return typeof v === 'number' ? v : null;
};

/**
 * Coalesces writes to one per animation frame, so dragging a slider recalculates at most once per
 * frame (WI-2: debounce ≤ 50 ms) and always ends on the last value.
 */
function useFrameWriter(onWrite: (entries: CellEntry[]) => string | null) {
  const pending = useRef<CellEntry | null>(null);
  const frame = useRef(0);
  useEffect(() => () => cancelAnimationFrame(frame.current), []);
  return (entry: CellEntry) => {
    pending.current = entry;
    if (frame.current) return;
    frame.current = requestAnimationFrame(() => {
      frame.current = 0;
      const next = pending.current;
      pending.current = null;
      if (next) onWrite([next]);
    });
  };
}

function Slider({ control, store, onWrite, context }: ControlProps<SliderControl>) {
  const ref = parseRef(control.cell);
  const id = useId();
  const write = useFrameWriter(onWrite);
  const value = numberAt(store, ref);
  const base = store.baseValue(ref);
  const label = labelText(store, controlLabelCell(control));
  const fullLabel = context ? `${context}: ${label}` : label;
  return (
    <div className={styles.row}>
      <label htmlFor={id} className={styles.label}>
        {context && <span className={styles.srOnly}>{context}: </span>}
        {label}
      </label>
      <div className={styles.sliderLine}>
        <input
          id={id}
          type="range"
          min={control.min}
          max={control.max}
          step={control.step}
          value={value ?? control.min}
          list={`${id}-ticks`}
          aria-valuetext={formatCell(value, control.fmt)}
          onChange={(e) => write({ ref, content: Number(e.target.value) })}
        />
        <datalist id={`${id}-ticks`}>
          {typeof base === 'number' && <option value={base} />}
        </datalist>
        <ValueBox
          cell={ref}
          store={store}
          fmt={control.fmt}
          label={`${fullLabel} value`}
          onWrite={onWrite}
        />
      </div>
      <BaseNote store={store} cell={ref} fmt={control.fmt} />
    </div>
  );
}

/** The slider's typed value box: shows the formatted value, accepts "5%", "0.05", "1,000". */
function ValueBox({
  cell,
  store,
  fmt,
  label,
  onWrite,
}: {
  cell: CellRef;
  store: WorkbookStore;
  fmt: string;
  label: string;
  onWrite: (entries: CellEntry[]) => string | null;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const commit = () => {
    if (draft === null) return;
    const content = parseInput(draft);
    setDraft(null);
    if (typeof content === 'number') onWrite([{ ref: cell, content }]);
  };
  return (
    <input
      className={styles.valueBox}
      aria-label={label}
      value={draft ?? formatCell(store.getValue(cell), fmt)}
      inputMode="decimal"
      spellCheck={false}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === 'Enter') commit();
        else if (e.key === 'Escape') setDraft(null);
      }}
    />
  );
}

function BaseNote({ store, cell, fmt }: { store: WorkbookStore; cell: CellRef; fmt?: string }) {
  const base = store.baseValue(cell);
  const now = store.getValue(cell);
  if (base === now) return null;
  return <span className={styles.base}>Base: {formatCell(base, fmt ?? store.fmt(cell))}</span>;
}

/** A number box that writes on every valid keystroke; the arrows step by `step`. */
function NumberInput({
  cell,
  store,
  min,
  max,
  step,
  label,
  id,
  onWrite,
}: {
  cell: CellRef;
  store: WorkbookStore;
  min?: number;
  max?: number;
  step: number;
  label?: string;
  id?: string;
  onWrite: (entries: CellEntry[]) => string | null;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const value = numberAt(store, cell);
  return (
    <input
      id={id}
      className={styles.number}
      type="number"
      aria-label={label}
      min={min}
      max={max}
      step={step}
      value={draft ?? (value === null ? '' : toSignificant15(value))}
      onChange={(e) => {
        setDraft(e.target.value);
        const n = e.target.valueAsNumber;
        if (e.target.value !== '' && Number.isFinite(n)) onWrite([{ ref: cell, content: n }]);
      }}
      onBlur={() => setDraft(null)}
    />
  );
}

function NumberRow({ control, store, onWrite }: ControlProps<NumberControl>) {
  const ref = parseRef(control.cell);
  const id = useId();
  return (
    <div className={styles.row}>
      <label htmlFor={id} className={styles.label}>
        {labelText(store, controlLabelCell(control))}
      </label>
      <NumberInput
        id={id}
        cell={ref}
        store={store}
        min={control.min}
        max={control.max}
        step={control.step}
        onWrite={onWrite}
      />
      <BaseNote store={store} cell={ref} fmt={control.fmt} />
    </div>
  );
}

function Select({ control, store, onWrite }: ControlProps<SelectControl>) {
  const ref = parseRef(control.cell);
  const id = useId();
  const value = formatCell(store.getValue(ref));
  const known = control.options.includes(value);
  return (
    <div className={styles.row}>
      <label htmlFor={id} className={styles.label}>
        {labelText(store, controlLabelCell(control))}
      </label>
      <select
        id={id}
        className={styles.select}
        value={value}
        onChange={(e) => onWrite([{ ref, content: e.target.value }])}
      >
        {/* A value typed into the grid that isn't an option still shows. */}
        {!known && (
          <option value={value} disabled>
            {value || '(blank)'}
          </option>
        )}
        {control.options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
      <BaseNote store={store} cell={ref} />
    </div>
  );
}

function Segmented({ control, store, onWrite }: ControlProps<SegmentedControl>) {
  const ref = parseRef(control.cell);
  const name = useId();
  const value = store.getValue(ref);
  return (
    <fieldset className={`${styles.row} ${styles.segmentedRow}`}>
      <legend className={styles.label}>{labelText(store, controlLabelCell(control))}</legend>
      <div className={styles.segmented}>
        {control.options.map((option) => (
          <label key={option.value}>
            <input
              type="radio"
              name={name}
              value={option.value}
              checked={value === option.value}
              onChange={() => onWrite([{ ref, content: option.value }])}
            />
            <span>{option.label}</span>
          </label>
        ))}
      </div>
    </fieldset>
  );
}

function DateRow({ control, store, onWrite }: ControlProps<DateControl>) {
  const ref = parseRef(control.cell);
  const id = useId();
  const value = numberAt(store, ref);
  return (
    <div className={styles.row}>
      <label htmlFor={id} className={styles.label}>
        {labelText(store, controlLabelCell(control))}
      </label>
      <input
        id={id}
        className={styles.date}
        type="date"
        min={control.min}
        max={control.max}
        value={value === null ? '' : serialToIsoDate(value)}
        onChange={(e) => {
          const serial = isoDateToSerial(e.target.value);
          if (serial !== undefined) onWrite([{ ref, content: serial }]);
        }}
      />
      <BaseNote store={store} cell={ref} fmt="yyyy-mm-dd" />
    </div>
  );
}

function Table({ control, store, onWrite }: ControlProps<TableControl>) {
  const cells = tableCells(control);
  const labels = tableLabels(control);
  const caption = labels.cols.length === 1 ? labelText(store, labels.cols[0]!) : null;
  const table = (
    <table className={styles.table}>
      {labels.cols.length > 1 && (
        <thead>
          <tr>
            <td />
            {labels.cols.map((c) => (
              <th key={c.cell} scope="col">
                {labelText(store, c)}
              </th>
            ))}
          </tr>
        </thead>
      )}
      <tbody>
        {cells.map((row, r) => {
          const rowLabel = formatCell(store.getValue(labels.rows[r]!), store.fmt(labels.rows[r]!));
          return (
            <tr key={r}>
              <th scope="row">{rowLabel}</th>
              {row.map((cell, c) => {
                const column = control.columns[c]!;
                const colLabel = labels.cols[c] ? labelText(store, labels.cols[c]) : '';
                return (
                  <td key={cell.cell}>
                    <NumberInput
                      cell={cell}
                      store={store}
                      min={column.min}
                      step={column.step}
                      label={`${rowLabel} ${colLabel} (${cell.sheet}!${cell.cell})`}
                      onWrite={onWrite}
                    />
                  </td>
                );
              })}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
  const summary = `${caption ?? 'Inputs'} (${control.range})`;
  return control.collapsed ? (
    <details className={styles.row}>
      <summary className={styles.label}>{summary}</summary>
      <div className={styles.tableScroll}>{table}</div>
    </details>
  ) : (
    <div className={styles.row}>
      <span className={styles.label}>{summary}</span>
      <div className={styles.tableScroll}>{table}</div>
    </div>
  );
}

/** WI-3: applies an experiment's changes; a second press restores what the cells held before. */
function ExperimentButton({
  control,
  store,
  experiments,
  onWrite,
}: ControlProps<ExperimentAction>) {
  const experiment = findExperiment(control.experiment, experiments);
  const changes = Object.entries(experiment.changes).map(([key, value]) => ({
    ref: parseRef(key),
    content: experimentChangeContent(value),
  }));
  const previous = useRef<CellEntry[] | null>(null);
  const active = changes.every(({ ref, content }) => store.getValue(ref) === content);
  return (
    <div className={styles.row}>
      <button
        type="button"
        className={styles.action}
        aria-pressed={active}
        onClick={() => {
          if (active) {
            const restore =
              previous.current ??
              changes.map(({ ref }) => ({ ref, content: store.originalContent(ref) }));
            if (onWrite(restore) === null) previous.current = null;
          } else {
            previous.current = changes.map(({ ref }) => {
              const v = store.getValue(ref);
              // Inputs hold constants; an error (from a formula typed in the grid) restores the original.
              return {
                ref,
                content: typeof v === 'object' && v !== null ? store.originalContent(ref) : v,
              };
            });
            onWrite(changes);
          }
        }}
      >
        {active ? control.undoLabel : control.label}
      </button>
      <span className={styles.base}>
        {experiment.id}: {experiment.title}
      </span>
    </div>
  );
}
