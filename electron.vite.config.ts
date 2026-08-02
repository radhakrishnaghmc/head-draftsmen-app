import { resolve } from 'path'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'electron/main.ts'),
          embeddingsWorker: resolve(__dirname, 'electron/embeddingsWorker.ts'),
          ocrWorker: resolve(__dirname, 'electron/ocrWorker.ts'),
          splitWorker: resolve(__dirname, 'electron/splitWorker.ts')
        }
      }
    },
    resolve: {
      alias: { '@core': resolve(__dirname, 'core') }
    }
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: { index: resolve(__dirname, 'electron/preload.ts') }
      }
    }
  },
  renderer: {
    root: resolve(__dirname, 'src'),
    // ES-module workers: the pdf.js worker shim (src/pdfWorkerShim.ts) applies
    // compat polyfills then dynamically imports the real pdf.worker, and that
    // code-split isn't allowed under the default IIFE worker format.
    worker: { format: 'es' },
    build: {
      rollupOptions: {
        input: { index: resolve(__dirname, 'src/index.html') }
      }
    },
    resolve: {
      alias: {
        '@core': resolve(__dirname, 'core'),
        '@': resolve(__dirname, 'src')
      }
    },
    plugins: [react()]
  }
})
