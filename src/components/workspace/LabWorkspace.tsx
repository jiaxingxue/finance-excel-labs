// The lab workspace (PRD §5.2): lesson, grid, checks, inspector, and in What-if mode (§6.4) the
// controls. The mode is the `?mode=` query parameter, so `/#/lab/7?mode=whatif` is a deep link.
// The lesson renders at once; the formula engine loads in the background (PRD §8).

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useSearchParams } from 'react-router';
import assertionsJson from '../../data/assertions.json' with { type: 'json' };
import exercisesJson from '../../data/exercises.json' with { type: 'json' };
import manifest from '../../data/manifest.json' with { type: 'json' };
import type { Assertion, Exercise, Lesson, LessonSummary } from '../../content/types.ts';
import { labs, type LabConfig } from '../../config/labs.config.ts';
import { toA1 } from '../../engine/address.ts';
import type { CellEntry, CellRef, RangeRef } from '../../engine/types.ts';
import { formulaBarText } from '../../format/displayValue.ts';
import { assertionPasses } from '../../grading/match.ts';
import { linkScopeForLab } from '../../lesson/cellRefs.ts';
import { loadEngineBundle } from '../../workspace/loadEngine.ts';
import { parseInput } from '../../workspace/parseInput.ts';
import {
  cellSelection,
  rangeSelection,
  selectionLabel,
  type Selection,
} from '../../workspace/selection.ts';
import { useWorkbookVersion } from '../../workspace/useWorkbook.ts';
import type { WorkbookStore } from '../../workspace/workbookStore.ts';
import { FormulaBar } from '../grid/FormulaBar.tsx';
import { Grid } from '../grid/Grid.tsx';
import { SheetTabs } from '../grid/SheetTabs.tsx';
import { LessonPanel } from '../lesson/LessonPanel.tsx';
import { VerificationNotice } from '../notice/VerificationNotice.tsx';
import { CellInspector } from '../side/CellInspector.tsx';
import { ChecksPanel, type CheckRow } from '../side/ChecksPanel.tsx';
import { WhatIfPanel } from '../whatif/WhatIfPanel.tsx';
import { LabHeader, type Mode } from './LabHeader.tsx';
import styles from './Workspace.module.css';

const assertions = assertionsJson as Assertion[];
const exercises = exercisesJson as Exercise[];
const allLabs = (manifest.lessons as LessonSummary[]).filter((l) => l.kind === 'lab');
// Every workbook sheet appears in some lab (tests/labs-config.test.ts), so the lesson can link
// sheet names before the engine chunk, which carries workbook.json, has loaded.
const allSheets = [...new Set(labs.flatMap((l) => l.sheets))];

type PhoneTab = 'lesson' | 'sheet' | 'checks';

interface LabWorkspaceProps {
  lab: LabConfig;
  summary: LessonSummary;
  /** Null while the lesson chunk loads. */
  lesson: Lesson | null;
}

export function LabWorkspace({ lab, summary, lesson }: LabWorkspaceProps) {
  const hasSheets = lab.sheets.length > 0;
  const [store, setStore] = useState<WorkbookStore | null>(null);
  const [engineError, setEngineError] = useState<string | null>(null);
  const [selection, setSelection] = useState<Selection>(() =>
    cellSelection(lab.sheets[0] ?? '', 'A1'),
  );
  const [showAll, setShowAll] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const [phoneTab, setPhoneTab] = useState<PhoneTab>('lesson');
  const [drawerOpen, setDrawerOpen] = useState(false);
  const lastOnSheet = useRef(new Map<string, Selection>());
  const bodyRef = useRef<HTMLDivElement>(null);
  const version = useWorkbookVersion(store);
  const [params, setParams] = useSearchParams();
  const whatIfAvailable = lab.controls.length > 0;
  const mode: Mode = params.get('mode') === 'whatif' && whatIfAvailable ? 'whatif' : 'explore';

  function setMode(next: Mode) {
    setParams(next === 'explore' ? {} : { mode: next }, { replace: true });
  }

  // Load the engine only after the lesson is on screen.
  const lessonReady = lesson !== null;
  useEffect(() => {
    if (!lessonReady || !hasSheets) return;
    let created: WorkbookStore | null = null;
    let live = true;
    loadEngineBundle()
      .then((bundle) => {
        if (!live) return;
        created = bundle.createStore();
        setStore(created);
      })
      .catch((error: unknown) => {
        if (live) setEngineError(error instanceof Error ? error.message : String(error));
      });
    return () => {
      live = false;
      created?.destroy();
    };
  }, [lessonReady, hasSheets]);

  const scope = useMemo(() => linkScopeForLab(lab, allSheets), [lab]);
  const visibleSheets = showAll && store ? store.sheetOrder : lab.sheets;
  const readOnly = !lab.sheets.includes(selection.sheet);
  const isReadOnly = (sheet: string) => !lab.sheets.includes(sheet);

  const focusGrid = () =>
    bodyRef.current?.querySelector<HTMLElement>('[role="grid"]')?.focus({ preventScroll: true });

  const select = useCallback((next: Selection) => {
    lastOnSheet.current.set(next.sheet, next);
    setSelection(next);
    setStatus(null);
  }, []);

  const onCellLink = useCallback(
    (range: RangeRef) => {
      if (!lab.sheets.includes(range.sheet)) setShowAll(true);
      select(rangeSelection(range));
      setPhoneTab('sheet');
      setDrawerOpen(false);
      // After the grid has rendered the selection, bring it into view and give it focus.
      requestAnimationFrame(() => {
        const grid = bodyRef.current?.querySelector<HTMLElement>('[role="grid"]');
        grid?.focus({ preventScroll: true });
        grid?.closest('section')?.scrollIntoView({ block: 'nearest' });
      });
    },
    [lab, select],
  );

  function selectSheet(sheet: string) {
    select(lastOnSheet.current.get(sheet) ?? cellSelection(sheet, 'A1'));
  }

  function switchSheet(delta: 1 | -1) {
    const i = visibleSheets.indexOf(selection.sheet);
    const next = visibleSheets[(i + delta + visibleSheets.length) % visibleSheets.length];
    if (next) selectSheet(next);
  }

  /** Commits typed text to a cell, which need not be the selected one (formula bar on blur). */
  function commit(ref: CellRef, typed: string): boolean {
    if (!store) return false;
    if (isReadOnly(ref.sheet)) {
      blockedEdit(ref.sheet);
      return false;
    }
    const error = store.enter(ref, parseInput(typed));
    setStatus(error ? `That entry wasn’t accepted: ${error}` : null);
    return error === null;
  }

  function write(entries: CellEntry[]): string | null {
    if (!store) return 'The formula engine is still loading.';
    const error = store.enterMany(entries);
    setStatus(error ? `That value wasn’t accepted: ${error}` : null);
    return error;
  }

  function blockedEdit(sheet = selection.sheet) {
    setStatus(`${sheet} belongs to another lab, so it’s read-only here.`);
  }

  function selectCell({ sheet, cell }: CellRef) {
    if (!visibleSheets.includes(sheet)) setShowAll(true);
    select(cellSelection(sheet, cell));
    setPhoneTab('sheet');
  }

  function resetLab(s: WorkbookStore) {
    s.reset();
    setStatus(null);
  }

  function onShowAll(next: boolean) {
    setShowAll(next);
    if (!next && !lab.sheets.includes(selection.sheet)) selectSheet(lab.sheets[0]!);
  }

  const checkRows = useMemo<CheckRow[] | null>(() => {
    if (!store) return null;
    void version; // recompute after every recalculation
    return assertions
      .filter((a) => lab.primarySheets.includes(a.sheet))
      .map((assertion) => {
        const actual = store.getValue(assertion);
        return {
          assertion,
          actual,
          pass: assertionPasses(actual, assertion),
          fmt: store.fmt(assertion),
        };
      });
  }, [store, version, lab]);

  const active = { sheet: selection.sheet, cell: toA1(selection.anchor) };
  const activeCell = store
    ? { value: store.getValue(active), formula: store.getFormula(active), fmt: store.fmt(active) }
    : null;
  const exercise = exercises.find(
    (e) => e.lab === lab.n && e.sheet === active.sheet && e.cell === active.cell,
  );
  const progress =
    checkRows && checkRows.length > 0
      ? { passed: checkRows.filter((r) => r.pass).length, total: checkRows.length }
      : null;

  const lessonPanel = (
    <section
      className={styles.lesson}
      aria-label="Lesson"
      data-drawer={drawerOpen ? 'open' : 'closed'}
    >
      {hasSheets && (
        <button
          type="button"
          className={`${styles.button} ${styles.drawerClose}`}
          onClick={() => setDrawerOpen(false)}
        >
          Close lesson
        </button>
      )}
      {lesson ? (
        <LessonPanel markdown={lesson.markdown} scope={scope} onCellLink={onCellLink} />
      ) : (
        <p className={styles.muted}>Loading the lesson…</p>
      )}
    </section>
  );

  return (
    <div className={styles.workspace}>
      <LabHeader
        lab={summary}
        allLabs={allLabs}
        progress={progress}
        onReset={store ? () => resetLab(store) : null}
        onToggleLesson={hasSheets ? () => setDrawerOpen((o) => !o) : null}
        mode={mode}
        whatIfAvailable={whatIfAvailable}
        onMode={setMode}
      />
      <VerificationNotice />

      {!hasSheets ? (
        <div className={styles.lessonOnly}>
          <p className={styles.note} role="note">
            Lab 11 has no workbook sheets. Its sandbox for the simulated function library arrives in
            a later milestone; for now, read the lesson below.
          </p>
          {lessonPanel}
        </div>
      ) : (
        <>
          <div className={styles.phoneTabs} role="group" aria-label="Workspace panels">
            {(['lesson', 'sheet', 'checks'] as const).map((tab) => (
              <button
                key={tab}
                type="button"
                aria-pressed={phoneTab === tab}
                onClick={() => setPhoneTab(tab)}
              >
                {tab === 'lesson'
                  ? 'Lesson'
                  : tab === 'sheet'
                    ? 'Sheet'
                    : mode === 'whatif'
                      ? 'What-if'
                      : 'Checks'}
              </button>
            ))}
          </div>
          <div ref={bodyRef} className={styles.body} data-tab={phoneTab}>
            {lessonPanel}
            {drawerOpen && (
              <div
                className={styles.scrim}
                aria-hidden="true"
                onClick={() => setDrawerOpen(false)}
              />
            )}

            <section className={styles.sheet} aria-label="Spreadsheet">
              <FormulaBar
                label={selectionLabel(selection)}
                target={active}
                contents={
                  activeCell
                    ? formulaBarText(activeCell.value, activeCell.formula, activeCell.fmt)
                    : ''
                }
                readOnly={!store || readOnly}
                onCommit={commit}
                onDone={focusGrid}
              />
              {status && (
                <p className={styles.status} role="alert">
                  {status}
                </p>
              )}
              {engineError ? (
                <p className={styles.status} role="alert">
                  The formula engine failed to load: {engineError}. Reload the page to try again.
                </p>
              ) : store ? (
                <Grid
                  key={selection.sheet}
                  store={store}
                  version={version}
                  selection={selection}
                  onSelect={select}
                  onCommit={(cell, typed) => commit({ sheet: selection.sheet, cell }, typed)}
                  onBlockedEdit={blockedEdit}
                  readOnly={readOnly}
                  onSwitchSheet={switchSheet}
                />
              ) : (
                <div className={styles.loading} role="status">
                  Loading the formula engine…
                </div>
              )}
              <SheetTabs
                sheets={visibleSheets}
                labSheets={lab.sheets}
                active={selection.sheet}
                showAll={showAll}
                onShowAll={onShowAll}
                onSelect={selectSheet}
              />
            </section>

            <div className={styles.checks} data-mode={mode}>
              {mode === 'whatif' &&
                (store ? (
                  <WhatIfPanel
                    lab={lab}
                    store={store}
                    version={version}
                    onWrite={write}
                    onSelectCell={selectCell}
                  />
                ) : (
                  <p className={styles.muted}>Loading the formula engine…</p>
                ))}
              <ChecksPanel
                rows={checkRows}
                emptyNote={`Lab ${lab.n} owns no checks.`}
                onSelect={(sheet, cell) => selectCell({ sheet, cell })}
              />
            </div>
            <div className={styles.inspector}>
              <CellInspector
                label={`${active.sheet}!${active.cell}`}
                cell={activeCell}
                exercise={exercise}
              />
            </div>
          </div>
        </>
      )}
    </div>
  );
}
