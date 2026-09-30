// The spreadsheet grid (GR-1, GR-2, GR-4, GR-6, GR-12, GR-13). A plain <table role="grid">: sheets are at
// most 57 rows × 15 columns, so nothing is virtualized (PRD §9.1). The table holds keyboard focus
// and points at the active cell with aria-activedescendant. The parent keys it by sheet, so
// switching sheets drops an edit in progress.

import {
  memo,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
  type MouseEvent,
} from 'react';
import { labs } from '../../config/labs.config.ts';
import { toA1, type GridPosition } from '../../engine/address.ts';
import { isErrorValue } from '../../engine/types.ts';
import { formatCell, formulaBarText } from '../../format/displayValue.ts';
import {
  jump,
  normalize,
  selectionLabel,
  step,
  type Bounds,
  type Direction,
  type Selection,
} from '../../workspace/selection.ts';
import { sheetFormats, type CellFormat } from '../../workspace/conditionalFormat.ts';
import type { WorkbookStore } from '../../workspace/workbookStore.ts';
import styles from './Grid.module.css';

// Conditional formats belong to their sheet, as in Excel, so they show in every lab that shows it.
const formatRules = labs.flatMap((l) => l.conditionalFormats);
const NO_FORMAT: CellFormat = {};

interface GridProps {
  store: WorkbookStore;
  /** Store version; a new value re-renders the grid after recalculation. */
  version: number;
  selection: Selection;
  onSelect: (selection: Selection) => void;
  /** Commits typed text to a cell; returns false if the edit was rejected. */
  onCommit: (cell: string, typed: string) => boolean;
  /** Called instead of editing when the sheet is read-only. */
  onBlockedEdit: () => void;
  readOnly: boolean;
  onSwitchSheet: (delta: 1 | -1) => void;
}

interface Editing {
  pos: GridPosition;
  text: string;
}

const MIN_ROWS = 20;
const MIN_COLS = 8;
const PAGE = 10;
const ARROWS: Record<string, Direction> = {
  ArrowUp: 'up',
  ArrowDown: 'down',
  ArrowLeft: 'left',
  ArrowRight: 'right',
};

const cellId = (sheet: string, cell: string) => `cell-${sheet}-${cell}`;

export function Grid({
  store,
  version,
  selection,
  onSelect,
  onCommit,
  onBlockedEdit,
  readOnly,
  onSwitchSheet,
}: GridProps) {
  const { sheet } = selection;
  const tableRef = useRef<HTMLTableElement>(null);
  const dragging = useRef(false);
  const [editing, setEditingState] = useState<Editing | null>(null);
  // Mirrors `editing` so a commit can't run twice (e.g. Enter, then blur as the input unmounts).
  const editingRef = useRef<Editing | null>(null);
  const setEditing = (next: Editing | null) => {
    editingRef.current = next;
    setEditingState(next);
  };

  const extent = store.extent(sheet);
  // Conditional formats for this sheet; the engine evaluates each rule.
  const formats = useMemo(() => {
    void version; // recompute after every recalculation
    return sheetFormats(formatRules, sheet, store);
  }, [store, sheet, version]);
  const sel = normalize(selection);
  const bounds: Bounds = {
    rows: Math.max(extent.rows + 3, sel.bottom + 2, MIN_ROWS),
    cols: Math.max(extent.cols + 2, sel.right + 2, MIN_COLS),
  };
  const activeId = cellId(sheet, toA1(selection.anchor));

  useEffect(() => {
    document.getElementById(activeId)?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }, [activeId]);

  const ref = (pos: GridPosition) => ({ sheet, cell: toA1(pos) });
  const isFilled = (pos: GridPosition) => store.getValue(ref(pos)) !== null;
  const select = (anchor: GridPosition, focus = anchor) => onSelect({ sheet, anchor, focus });
  const focusTable = () => tableRef.current?.focus({ preventScroll: true });

  function startEdit(text?: string) {
    if (readOnly) return onBlockedEdit();
    const pos = selection.anchor;
    const r = ref(pos);
    setEditing({
      pos,
      text: text ?? formulaBarText(store.getValue(r), store.getFormula(r), store.fmt(r)),
    });
  }

  function finishEdit(commit: boolean, move?: Direction) {
    const current = editingRef.current;
    if (!current) return;
    const { pos, text } = current;
    setEditing(null);
    if (commit && !onCommit(toA1(pos), text)) {
      focusTable();
      return;
    }
    if (move) select(step(pos, move, bounds));
    focusTable();
  }

  function onKeyDown(e: KeyboardEvent<HTMLTableElement>) {
    if (editing) return; // the editor's own handler runs
    const ctrl = e.ctrlKey || e.metaKey;
    const { anchor, focus } = selection;
    const dir = ARROWS[e.key];
    let handled = true;
    if (dir) {
      const from = e.shiftKey ? focus : anchor;
      const to = ctrl ? jump(from, dir, bounds, isFilled) : step(from, dir, bounds);
      if (e.shiftKey) select(anchor, to);
      else select(to);
    } else if (e.key === 'PageDown' || e.key === 'PageUp') {
      if (ctrl) onSwitchSheet(e.key === 'PageDown' ? 1 : -1);
      else select(step(anchor, e.key === 'PageDown' ? 'down' : 'up', bounds, PAGE));
    } else if (e.key === 'Home') {
      select(ctrl ? { row: 0, col: 0 } : { row: anchor.row, col: 0 });
    } else if (e.key === 'End') {
      const last = { row: Math.max(extent.rows - 1, 0), col: Math.max(extent.cols - 1, 0) };
      select(ctrl ? last : { row: anchor.row, col: last.col });
    } else if (e.key === 'Enter' || e.key === 'F2') {
      startEdit();
    } else if (e.key === 'Delete') {
      if (readOnly) onBlockedEdit();
      else onCommit(toA1(anchor), '');
    } else if (e.key === 'Backspace') {
      startEdit('');
    } else if (ctrl && e.key.toLowerCase() === 'z' && !e.shiftKey) {
      store.undo();
    } else if (
      ctrl &&
      (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))
    ) {
      store.redo();
    } else if (e.key.length === 1 && !ctrl && !e.altKey) {
      startEdit(e.key);
    } else {
      handled = false;
    }
    if (handled) e.preventDefault();
  }

  function positionOf(e: MouseEvent): GridPosition | null {
    const td = (e.target as HTMLElement).closest<HTMLElement>('td[data-row]');
    return td ? { row: Number(td.dataset.row), col: Number(td.dataset.col) } : null;
  }

  function onMouseDown(e: MouseEvent) {
    const pos = positionOf(e);
    if (!pos || (editing && editing.pos.row === pos.row && editing.pos.col === pos.col)) return;
    e.preventDefault(); // keep focus on the table, not the cell
    if (editing) finishEdit(true);
    dragging.current = true;
    if (e.shiftKey) select(selection.anchor, pos);
    else select(pos);
    focusTable();
  }

  function onMouseOver(e: MouseEvent) {
    showFullTextIfCut(e.target as HTMLElement);
    if (!dragging.current || e.buttons !== 1) return;
    const pos = positionOf(e);
    if (pos) select(selection.anchor, pos);
  }

  const label = selectionLabel(selection);
  const active = ref(selection.anchor);
  const activeFormula = store.getFormula(active);
  const announcement = `${label}, ${formatCell(store.getValue(active), store.fmt(active)) || 'blank'}${
    activeFormula ? `, formula ${activeFormula}` : ''
  }`;

  return (
    <div className={styles.scroller}>
      <table
        ref={tableRef}
        role="grid"
        aria-label={`Sheet ${sheet}`}
        aria-rowcount={bounds.rows + 1}
        aria-colcount={bounds.cols + 1}
        aria-activedescendant={activeId}
        aria-readonly={readOnly || undefined}
        tabIndex={0}
        className={styles.grid}
        onKeyDown={onKeyDown}
        onMouseDown={onMouseDown}
        onMouseOver={onMouseOver}
        onMouseUp={() => (dragging.current = false)}
        onDoubleClick={(e) => positionOf(e) && startEdit()}
      >
        <thead>
          <tr>
            <th className={styles.corner} aria-hidden="true" />
            {Array.from({ length: bounds.cols }, (_, col) => (
              <th
                key={col}
                scope="col"
                className={col >= sel.left && col <= sel.right ? styles.headActive : undefined}
              >
                {toA1({ row: 0, col }).replace(/\d+$/, '')}
              </th>
            ))}
          </tr>
        </thead>
        <tbody data-version={version}>
          {Array.from({ length: bounds.rows }, (_, row) => (
            <tr key={row}>
              <th
                scope="row"
                className={row >= sel.top && row <= sel.bottom ? styles.headActive : undefined}
              >
                {row + 1}
              </th>
              {Array.from({ length: bounds.cols }, (_, col) => {
                const cell = toA1({ row, col });
                const r = { sheet, cell };
                const isEditing = editing?.pos.row === row && editing.pos.col === col;
                if (isEditing) {
                  return (
                    <td
                      key={col}
                      id={cellId(sheet, cell)}
                      role="gridcell"
                      aria-selected
                      className={styles.editing}
                    >
                      <CellEditor
                        label={`Edit ${sheet}!${cell}`}
                        text={editing.text}
                        onChange={(text) => setEditing({ pos: editing.pos, text })}
                        onFinish={finishEdit}
                      />
                    </td>
                  );
                }
                const value = store.getValue(r);
                const inRange =
                  row >= sel.top && row <= sel.bottom && col >= sel.left && col <= sel.right;
                const key = `${sheet}!${cell}`;
                const format = formats.get(cell) ?? NO_FORMAT;
                return (
                  <GridCell
                    key={col}
                    id={cellId(sheet, cell)}
                    row={row}
                    col={col}
                    text={formatCell(value, store.fmt(r))}
                    kind={
                      value === null
                        ? 'blank'
                        : isErrorValue(value)
                          ? 'error'
                          : typeof value === 'number'
                            ? 'number'
                            : typeof value === 'boolean'
                              ? 'boolean'
                              : 'text'
                    }
                    input={value !== null && store.getFormula(r) === null}
                    active={row === selection.anchor.row && col === selection.anchor.col}
                    inRange={inRange}
                    flashSeq={store.lastChange.keys.has(key) ? store.lastChange.seq : 0}
                    fill={format.fill}
                    bold={format.bold}
                    scale={format.scale}
                  />
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      <p className={styles.srOnly} aria-live="polite">
        {announcement}
      </p>
    </div>
  );
}

interface GridCellProps {
  id: string;
  row: number;
  col: number;
  text: string;
  kind: 'blank' | 'number' | 'text' | 'boolean' | 'error';
  input: boolean;
  active: boolean;
  inRange: boolean;
  /** Non-zero when the cell changed in the last recalculation; a new value replays the flash. */
  flashSeq: number;
  /** Conditional format (GR-12). */
  fill?: CellFormat['fill'];
  bold?: boolean;
  scale?: number;
}

const GridCell = memo(function GridCell({
  id,
  row,
  col,
  text,
  kind,
  input,
  active,
  inRange,
  flashSeq,
  fill,
  bold,
  scale,
}: GridCellProps) {
  const className = [
    styles.cell,
    styles[kind],
    input && styles.input,
    fill && styles[fill],
    scale !== undefined && styles.scale,
    bold && styles.bold,
    inRange && styles.inRange,
    active && styles.active,
  ]
    .filter(Boolean)
    .join(' ');
  return (
    <td
      id={id}
      role="gridcell"
      aria-selected={inRange}
      data-row={row}
      data-col={col}
      className={className}
      style={
        scale === undefined
          ? undefined
          : ({ '--cf-scale': `${Math.round(scale * 100)}%` } as CSSProperties)
      }
    >
      <span key={flashSeq} className={flashSeq ? styles.flash : undefined}>
        {text}
      </span>
    </td>
  );
});

interface CellEditorProps {
  label: string;
  text: string;
  onChange: (text: string) => void;
  onFinish: (commit: boolean, move?: Direction) => void;
}

function CellEditor({ label, text, onChange, onFinish }: CellEditorProps) {
  const done = useRef(false);
  const finish = (commit: boolean, move?: Direction) => {
    if (done.current) return;
    done.current = true;
    onFinish(commit, move);
  };
  return (
    <input
      className={styles.editor}
      aria-label={label}
      value={text}
      autoFocus
      spellCheck={false}
      onChange={(e) => onChange(e.target.value)}
      onBlur={() => finish(true)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') finish(true, e.shiftKey ? 'up' : 'down');
        else if (e.key === 'Tab') finish(true, e.shiftKey ? 'left' : 'right');
        else if (e.key === 'Escape') finish(false);
        else return;
        e.preventDefault();
        e.stopPropagation();
      }}
    />
  );
}

/**
 * Text cut off by the column width gets a hover tooltip with the full text; text that fits gets
 * none. Keyboard and screen-reader users get the full text from the formula bar and the live region.
 */
function showFullTextIfCut(target: HTMLElement) {
  const span = target.closest('td')?.querySelector<HTMLElement>(':scope > span');
  if (!span) return;
  if (span.scrollWidth > span.clientWidth) span.title = span.textContent ?? '';
  else span.removeAttribute('title');
}
