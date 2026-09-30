// LS-4: the verification notice from the top of the Labs document, verbatim, plus the app's own
// note that it calculates with an Excel-compatible engine, not Microsoft Excel (PRD §10).

import Markdown from 'react-markdown';
import notice from '../../data/notice.json' with { type: 'json' };
import styles from './VerificationNotice.module.css';

export function VerificationNotice({ open = false }: { open?: boolean }) {
  return (
    <details className={styles.notice} open={open}>
      <summary>How these labs were verified</summary>
      <div className={styles.body}>
        <Markdown>{notice.markdown.replace(/^> ?/gm, '')}</Markdown>
        <p>
          <strong>In this app,</strong> formulas are calculated in your browser by HyperFormula, an
          Excel-compatible formula engine, not by Microsoft Excel.
        </p>
      </div>
    </details>
  );
}
