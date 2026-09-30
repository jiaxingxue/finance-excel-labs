// GR-3: the selected cell's address and raw contents. Numbers show 15 significant digits
// (OPEN_ISSUES #19). The contents box is also an editor, which is the easy way to edit on a phone.

import { useState } from 'react';
import styles from './SheetChrome.module.css';

interface FormulaBarProps {
  /** e.g. "BvA!C6" or "BvA!A5:H5" */
  label: string;
  contents: string;
  readOnly: boolean;
  onCommit: (typed: string) => boolean;
  /** Where focus goes after Enter or Esc. */
  onDone: () => void;
}

export function FormulaBar({ label, contents, readOnly, onCommit, onDone }: FormulaBarProps) {
  // `draft` is null until the learner types, so a new selection or a recalculation shows through.
  const [draft, setDraft] = useState<{ label: string; text: string } | null>(null);
  const text = draft?.label === label ? draft.text : contents;

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
        onChange={(e) => setDraft({ label, text: e.target.value })}
        onBlur={() => setDraft(null)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            if (draft && !readOnly && !onCommit(draft.text)) return;
            setDraft(null);
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
