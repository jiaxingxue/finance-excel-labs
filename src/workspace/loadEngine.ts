// Lazy loader for the engine chunk (engineBundle.ts). The download starts once and is shared.

type EngineBundle = typeof import('./engineBundle.ts');

let pending: Promise<EngineBundle> | null = null;

export function loadEngineBundle(): Promise<EngineBundle> {
  pending ??= import('./engineBundle.ts').catch((error: unknown) => {
    pending = null; // let a later visit retry, e.g. after a network blip
    throw error;
  });
  return pending;
}
