import { Link } from 'react-router';
import { useDocumentTitle } from '../components/useDocumentTitle.ts';
import styles from './Page.module.css';

export function NotFound() {
  useDocumentTitle('Page not found');
  return (
    <div className={styles.standalone}>
      <h1 className={styles.title}>Page not found</h1>
      <p>
        <Link to="/">Go to the home page</Link>
      </p>
    </div>
  );
}
