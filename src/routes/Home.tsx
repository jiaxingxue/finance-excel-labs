import manifest from '../data/manifest.json' with { type: 'json' };
import { useDocumentTitle } from '../components/useDocumentTitle.ts';
import styles from './Page.module.css';

const labs = manifest.lessons.filter((l) => l.kind === 'lab');
const { counts } = manifest;

export function Home() {
  useDocumentTitle(null);
  return (
    <>
      <h1 className={styles.title}>Excel Labs for Financial Analysis</h1>
      <p className={styles.lead}>
        Interactive, auto-graded Excel labs for FP&amp;A and accounting: budget vs. actual,
        price–volume–mix, bank reconciliation, AR aging, forecasting, and driver-based scenarios.
      </p>

      <p className={styles.notice} role="status">
        <strong>In development.</strong> This placeholder confirms the build and content pipeline.
        The interactive labs arrive in later milestones.
      </p>

      <dl className={styles.stats} aria-label="Content loaded from the Labs document">
        <div>
          <dt>Expected values</dt>
          <dd>{counts.assertions}</dd>
        </div>
        <div>
          <dt>What-if experiments</dt>
          <dd>{counts.experiments}</dd>
        </div>
        <div>
          <dt>Sheets</dt>
          <dd>{counts.sheets}</dd>
        </div>
        <div>
          <dt>Formulas</dt>
          <dd>{counts.formulas}</dd>
        </div>
      </dl>

      <h2 className={styles.heading}>Labs</h2>
      <ol className={styles.labList}>
        {labs.map((lab) => (
          <li key={lab.id}>
            <span className={styles.labNumber}>Lab {lab.n}</span> {lab.title}
          </li>
        ))}
      </ol>
    </>
  );
}
