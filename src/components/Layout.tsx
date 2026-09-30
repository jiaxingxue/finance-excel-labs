import { NavLink, Outlet, useMatch } from 'react-router';
import styles from './Layout.module.css';

const REPO_URL = 'https://github.com/jiaxingxue/finance-excel-labs';

export function Layout() {
  const commit = import.meta.env.VITE_COMMIT_SHA;
  // The lab workspace uses the full window width; other pages keep a readable column.
  const wide = useMatch('/lab/:n') !== null;
  return (
    <div className={styles.shell}>
      <a className={styles.skip} href="#main">
        Skip to content
      </a>
      <header className={styles.header}>
        <NavLink to="/" className={styles.brand}>
          Excel Labs for Financial Analysis
        </NavLink>
        <nav aria-label="Main">
          <NavLink to="/about" className={styles.navLink}>
            About
          </NavLink>
        </nav>
      </header>
      <main
        id="main"
        className={wide ? `${styles.main} ${styles.wide}` : styles.main}
        tabIndex={-1}
      >
        <Outlet />
      </main>
      <footer className={styles.footer}>
        <span>
          <a href={REPO_URL}>@jiaxingxue on GitHub</a>
        </span>
        <span>
          <a href={`${REPO_URL}/blob/main/LICENSE`}>GPLv3</a>
        </span>
        <span>
          v{import.meta.env.VITE_APP_VERSION} ·{' '}
          <code title={commit}>{commit === 'dev' ? 'dev' : commit.slice(0, 7)}</code>
        </span>
      </footer>
    </div>
  );
}
