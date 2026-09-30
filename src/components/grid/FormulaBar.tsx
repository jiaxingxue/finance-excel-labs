// GR-3: the selected cell's address and raw contents. Numbers show 15 significant digits
// (OPEN_ISSUES #19). The contents box is also an editor, which is the easy way to edit on a phone.
// As in Excel, an unsaved edit commits when focus leaves the box, to the cell it was typed for,
// even if a click on the grid has already selected another cell. Esc discards it.

import { useRef, useState } from 'react';
import type { CellRef } from '../../engine/types.ts';
import styles from './SheetChrome.module.css';

interface FormulaBarProps {
  /** e.g. "BvA!C6" or "BvA!A5:H5" */
  label: string;
  /** The active cell, which an edit writes to. */
  target: CellRef;
  contents: string;
  readOnly: boolean;
  onCommit: (target: CellRef, typed: string) => boolean;
  /** Where focus goes after Enter or Esc. */
  onDone: () => void;
}

interface Draft {
  target: CellRef;
  text: string;
}

const sameCell = (a: CellRef, b: CellRef) => a.sheet === b.sheet && a.cell === b.cell;

export function FormulaBar({
  label,
  target,
  contents,
  readOnly,
  onCommit,
  onDone,
}: FormulaBarProps) {
  // `draft` is null until the learner types, so a new selection or a recalculation shows through.
  const [draft, setDraftState] = useState<Draft | null>(null);
  // Mirrors `draft` so Enter followed by the blur it causes can't commit twice.
  const draftRef = useRef<Draft | null>(null);
  const setDraft = (next: Draft | null) => {
    draftRef.current = next;
    setDraftState(next);
  };
  const text = draft && sameCell(draft.target, target) ? draft.text : contents;

  /** Commits a pending edit. Returns false if the entry was rejected. */
  const commit = (): boolean => {
    const pending = draftRef.current;
    setDraft(null);
    if (!pending || readOnly) return true;
    return onCommit(pending.target, pending.text);
  };

  return (
    <div className={styles.formulaBar}>
      <output className={styles.nameBox} aria-label="Selected cell">
        {label}
      </output>
      <span className={styles.fx} aria-hidden="true">
        fx
      </span>
      <input
        className={styles.contents}
        aria-label="Formula bar"
        value={text}
        readOnly={readOnly}
        spellCheck={false}
        autoComplete="off"
        onChange={(e) => setDraft({ target: draft?.target ?? target, text: e.target.value })}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            const pending = draftRef.current;
            if (!commit()) {
              setDraft(pending); // keep the text so it can be fixed
              return;
            }
            onDone();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setDraft(null);
            onDone();
          }
        }}
      />
    </div>
  );
}
