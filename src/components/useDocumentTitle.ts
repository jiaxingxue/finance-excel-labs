import { useEffect } from 'react';

const APP_NAME = 'Excel Labs';

/** Per-route page title, e.g. "About · Excel Labs" (PRD §17.11). */
export function useDocumentTitle(title: string | null): void {
  useEffect(() => {
    document.title = title ? `${title} · ${APP_NAME}` : 'Excel Labs for Financial Analysis';
  }, [title]);
}
