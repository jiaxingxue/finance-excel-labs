import { useDocumentTitle } from '../components/useDocumentTitle.ts';
import styles from './Page.module.css';

/** Last-resort screen for render errors. The full error boundary with "Reset this lab" is M7. */
export function RouteError() {
  useDocumentTitle('Something went wrong');
  return (
    <div className={styles.standalone}>
      <h1 className={styles.title}>Something went wrong</h1>
      <p>
        <a href={import.meta.env.BASE_URL}>Reload the home page</a> or{' '}
        <a href="https://github.com/jiaxingxue/finance-excel-labs/issues">report an issue</a>.
      </p>
    </div>
  );
}
