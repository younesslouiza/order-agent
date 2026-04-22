import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      external: [],
    },
    commonjsOptions: {
      transformMixedEsModules: true,
    }
  },
  resolve: {
    alias: {}
  },
  optimizeDeps: {
    include: ['xlsx'],
    esbuildOptions: {
      target: 'es2020'
    }
  }
})