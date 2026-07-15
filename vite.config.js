import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// 폐쇄망 대응: 외부 CDN 없이 전부 번들에 포함.
// base: './' 로 두면 file:// 또는 임의 하위경로에서도 dist 를 열 수 있음.
export default defineConfig({
  base: './',
  plugins: [react()],
  build: {
    outDir: 'dist',
    assetsInlineLimit: 0,
    chunkSizeWarningLimit: 4000,
  },
})
