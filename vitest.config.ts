import { defineConfig } from 'vitest/config'
import { resolve } from 'path'

export default defineConfig({
  resolve: {
    alias: { '@core': resolve(__dirname, 'core') }
  },
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts']
  }
})
