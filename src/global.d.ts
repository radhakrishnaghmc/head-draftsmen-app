import type { DocuGenApi } from '../electron/ipc-contract'

declare global {
  interface Window {
    docugen: DocuGenApi
  }
}

export {}
