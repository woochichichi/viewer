import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { viteSingleFile } from 'vite-plugin-singlefile'
import { readFileSync } from 'node:fs'

const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url)))

// SINGLE=1 로 빌드하면 모든 JS/CSS 를 하나의 HTML 안에 인라인 →
// "더블클릭 실행" 가능한 단일 파일 산출 (release/ 로 출력).
const single = process.env.SINGLE === '1'

// 폐쇄망 대응: 외부 CDN 없이 전부 번들에 포함.
// base: './' 로 두면 file:// 또는 임의 하위경로에서도 열 수 있음.
export default defineConfig({
  base: './',
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
  },
  plugins: [react(), ...(single ? [viteSingleFile()] : [])],
  build: single
    ? {
        outDir: 'release',
        assetsInlineLimit: 100000000,
        chunkSizeWarningLimit: 100000,
        cssCodeSplit: false,
        emptyOutDir: false,
      }
    : {
        outDir: 'dist',
        assetsInlineLimit: 0,
        chunkSizeWarningLimit: 4000,
      },
})
