import { lazy, Suspense } from 'react';

// The lab workspace (grid, markdown, and later the engine) loads only when a lab is opened.
const Lab = lazy(() => import('./Lab.tsx'));

export function LabRoute() {
  return (
    <Suspense fallback={<p role="status">Loading the lab…</p>}>
      <Lab />
    </Suspense>
  );
}
