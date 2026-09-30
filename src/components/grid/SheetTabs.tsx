// GR-1: tabs for the lab's sheets, plus a toggle that reveals every sheet (read-only outside the lab).

import styles from './SheetChrome.module.css';

interface SheetTabsProps {
  sheets: readonly string[];
  labSheets: readonly string[];
  active: string;
  showAll: boolean;
  onShowAll: (showAll: boolean) => void;
  onSelect: (sheet: string) => void;
}

export function SheetTabs({
  sheets,
  labSheets,
  active,
  showAll,
  onShowAll,
  onSelect,
}: SheetTabsProps) {
  return (
    <div className={styles.tabsRow}>
      <div className={styles.tabs} role="group" aria-label="Sheets">
        {sheets.map((sheet) => {
          const inLab = labSheets.includes(sheet);
          return (
            <button
              key={sheet}
              type="button"
              className={styles.tab}
              aria-pressed={sheet === active}
              title={inLab ? undefined : 'Another lab’s sheet (read-only here)'}
              onClick={() => onSelect(sheet)}
            >
              {sheet}
              {!inLab && <span className={styles.readOnlyMark}> (read-only)</span>}
            </button>
          );
        })}
      </div>
      <label className={styles.showAll}>
        <input type="checkbox" checked={showAll} onChange={(e) => onShowAll(e.target.checked)} />
        All sheets
      </label>
    </div>
  );
}
