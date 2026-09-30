import { useDocumentTitle } from '../components/useDocumentTitle.ts';
import styles from './Page.module.css';

export function About() {
  useDocumentTitle('About');
  const env = import.meta.env;
  return (
    <>
      <h1 className={styles.title}>About</h1>
      <p className={styles.lead}>
        A browser-based learning app for FP&amp;A and accounting. Formulas will run in an
        Excel-compatible engine in your browser, not in Microsoft Excel. This project has no
        affiliation with Microsoft.
      </p>

      <h2 className={styles.heading}>License</h2>
      <p>
        Code and lesson content are licensed under the GNU General Public License v3. See the{' '}
        <a href="https://github.com/jiaxingxue/finance-excel-labs/blob/main/LICENSE">LICENSE</a> and{' '}
        <a href="https://github.com/jiaxingxue/finance-excel-labs/blob/main/THIRD_PARTY_NOTICES.md">
          third-party notices
        </a>
        .
      </p>

      <h2 className={styles.heading}>Further reading</h2>
      <p>
        Two companion study guides (financial analysis concepts; CFO controls and COSO) are coming
        soon.
      </p>

      <h2 className={styles.heading}>Build</h2>
      <dl className={styles.meta}>
        <dt>Version</dt>
        <dd>{env.VITE_APP_VERSION}</dd>
        <dt>Commit</dt>
        <dd>
          <code>{env.VITE_COMMIT_SHA}</code>
        </dd>
        <dt>Built</dt>
        <dd>{env.VITE_BUILD_DATE}</dd>
      </dl>
    </>
  );
}
