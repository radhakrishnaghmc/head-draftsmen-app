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
          embeddingsWorker: resolve(__dirname, 'electron/embeddingsWorker.ts')
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
