import { useSyncExternalStore } from 'react';
import type { WorkbookStore } from './workbookStore.ts';

/** Re-renders the caller after every recalculation. Returns the store's version number. */
export function useWorkbookVersion(store: WorkbookStore | null): number {
  return useSyncExternalStore(store?.subscribe ?? noopSubscribe, store?.getVersion ?? zero);
}

const noopSubscribe = () => () => {};
const zero = () => 0;
