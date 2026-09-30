// LS-1 + GR-11: the lab's lesson, verbatim from the Labs document, rendered as GitHub-flavored
// markdown. Cell references become buttons that select the cell in the grid.

import { useMemo, type ReactNode } from 'react';
import Markdown, { type Components } from 'react-markdown';
import remarkGfm from 'remark-gfm';
import type { RangeRef } from '../../engine/types.ts';
import { parseRange, type LinkScope } from '../../lesson/cellRefs.ts';
import { highlightFormula } from '../../lesson/highlightFormula.ts';
import { rehypeCellLinks } from '../../lesson/rehypeCellLinks.ts';
import styles from './Lesson.module.css';

interface LessonPanelProps {
  markdown: string;
  scope: LinkScope;
  onCellLink: (range: RangeRef) => void;
}

export function LessonPanel({ markdown, scope, onCellLink }: LessonPanelProps) {
  // The workspace header already shows the "## Lab N — Title" heading as the page's h1.
  const body = useMemo(() => markdown.replace(/^## .*\n/, ''), [markdown]);

  const components = useMemo<Components>(
    () => ({
      a({ node, href, children }) {
        const ref = node?.properties?.dataCellRef;
        const range = typeof ref === 'string' ? parseRange(ref) : null;
        if (range) {
          return (
            <button
              type="button"
              className={styles.cellLink}
              title={`Select ${ref as string}`}
              onClick={() => onCellLink(range)}
            >
              {children}
            </button>
          );
        }
        // In-document anchors point at Labs-document sections this app doesn't render.
        if (!href || href.startsWith('#')) return <span>{children}</span>;
        return (
          <a href={href} target="_blank" rel="noreferrer">
            {children}
          </a>
        );
      },
      code({ className, children }) {
        if (className?.includes('language-excel') && typeof children === 'string') {
          return <code className={className}>{highlight(children)}</code>;
        }
        return <code className={className}>{children}</code>;
      },
      // Wide tables and code blocks scroll sideways; tabIndex lets keyboard users scroll them.
      pre({ children }) {
        return <pre tabIndex={0}>{children}</pre>;
      },
      table({ children }) {
        return (
          <div className={styles.tableScroll} tabIndex={0}>
            <table>{children}</table>
          </div>
        );
      },
    }),
    [onCellLink],
  );

  return (
    <div className={styles.lesson}>
      <Markdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[[rehypeCellLinks, scope]]}
        components={components}
      >
        {body}
      </Markdown>
    </div>
  );
}

function highlight(source: string): ReactNode[] {
  return highlightFormula(source).map((t, i) =>
    t.kind === 'text' ? (
      t.text
    ) : (
      <span key={i} className={styles[t.kind]}>
        {t.text}
      </span>
    ),
  );
}
