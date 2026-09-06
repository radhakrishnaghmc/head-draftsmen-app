/// <reference types="vite/client" />

interface ImportMetaEnv {
  // Local-only dev auto-login password (see src/components/LoginPage.tsx);
  // set in a gitignored .env(.local), never present in a real build.
  readonly VITE_DEV_PASSWORD?: string
}
