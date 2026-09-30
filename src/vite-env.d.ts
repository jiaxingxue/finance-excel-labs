/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** Full commit SHA, injected at build time (PRD §17.3). "dev" outside git. */
  readonly VITE_COMMIT_SHA: string;
  /** ISO timestamp of the build. */
  readonly VITE_BUILD_DATE: string;
  /** package.json version. */
  readonly VITE_APP_VERSION: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
