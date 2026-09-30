// Route /lab/:n (PRD §5.1). This module is lazy-loaded; each lesson is its own small chunk, and
// the formula engine is a third chunk that LabWorkspace requests once the lesson is on screen.

import { useEffect, useState } from 'react';
import { useParams } from 'react-router';
import manifest from '../data/manifest.json' with { type: 'json' };
import type { Lesson, LessonSummary } from '../content/types.ts';
import { labs } from '../config/labs.config.ts';
import { LabWorkspace } from '../components/workspace/LabWorkspace.tsx';
import { useDocumentTitle } from '../components/useDocumentTitle.ts';
import { NotFound } from './NotFound.tsx';

const lessonLoaders = import.meta.glob<Lesson>('../data/lessons/lab-*.json', { import: 'default' });

export default function Lab() {
  const { n } = useParams();
  const lab = labs.find((l) => String(l.n) === n);
  const summary = (manifest.lessons as LessonSummary[]).find(
    (l) => l.kind === 'lab' && String(l.n) === n,
  );
  const [lesson, setLesson] = useState<Lesson | null>(null);
  useDocumentTitle(summary ? `Lab ${summary.n} — ${summary.title}` : 'Page not found');

  useEffect(() => {
    if (!summary) return;
    let live = true;
    void lessonLoaders[`../data/lessons/${summary.id}.json`]?.().then((l) => live && setLesson(l));
    return () => {
      live = false;
    };
  }, [summary]);

  if (!lab || !summary) return <NotFound />;
  // Keyed by lab so switching labs starts a fresh workspace and engine.
  return (
    <LabWorkspace
      key={lab.n}
      lab={lab}
      summary={summary}
      lesson={lesson?.id === summary.id ? lesson : null}
    />
  );
}
