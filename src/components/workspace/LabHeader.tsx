// Top bar of the lab workspace (PRD §5.2): lab selector, mode tabs, progress, reset.

import { useNavigate } from 'react-router';
import type { LessonSummary } from '../../content/types.ts';
import styles from './Workspace.module.css';

/** Modes that later milestones add (PRD §12). */
const LATER_MODES = ['Build', 'What-if', 'Challenge', 'Modern'];

interface LabHeaderProps {
  lab: LessonSummary;
  allLabs: LessonSummary[];
  /** Null while the engine loads or when the lab has no checks. */
  progress: { passed: number; total: number } | null;
  onReset: (() => void) | null;
  onToggleLesson: (() => void) | null;
}

export function LabHeader({ lab, allLabs, progress, onReset, onToggleLesson }: LabHeaderProps) {
  const navigate = useNavigate();
  return (
    <div className={styles.header}>
      <div className={styles.titleRow}>
        <h1 className={styles.title}>
          Lab {lab.n} — {lab.title}
        </h1>
        <label className={styles.labPicker}>
          <span className={styles.srOnly}>Switch lab</span>
          <select value={lab.n ?? ''} onChange={(e) => navigate(`/lab/${e.target.value}`)}>
            {allLabs.map((l) => (
              <option key={l.id} value={l.n ?? ''}>
                Lab {l.n}
              </option>
            ))}
          </select>
        </label>
      </div>
      <div className={styles.toolbar}>
        <div className={styles.modes} role="group" aria-label="Modes">
          <button type="button" className={styles.mode} aria-pressed="true">
            Explore
          </button>
          {LATER_MODES.map((mode) => (
            <button key={mode} type="button" className={styles.mode} disabled title="Coming soon">
              {mode} <span className={styles.soon}>coming soon</span>
            </button>
          ))}
        </div>
        <div className={styles.actions}>
          {progress && (
            <span className={styles.progress} role="status">
              Checks {progress.passed}/{progress.total}{' '}
              <span aria-hidden="true">{progress.passed === progress.total ? '✓' : '✗'}</span>
            </span>
          )}
          {onToggleLesson && (
            <button
              type="button"
              className={`${styles.button} ${styles.lessonToggle}`}
              onClick={onToggleLesson}
            >
              Lesson
            </button>
          )}
          {onReset && (
            <button type="button" className={styles.button} onClick={onReset}>
              Reset lab
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
