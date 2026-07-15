// 빌드 후처리: 사용하지 않는 언어팩(.pak)을 제거해 용량을 줄인다.
// Electron 은 수십 개 언어팩을 포함하는데, 한국어/영어만 남긴다.
const fs = require('node:fs')
const path = require('node:path')

exports.default = async function afterPack(context) {
  const keep = new Set(['en-US.pak', 'ko.pak'])
  const localesDir = path.join(context.appOutDir, 'locales')
  try {
    const files = fs.readdirSync(localesDir)
    let removed = 0
    for (const f of files) {
      if (!keep.has(f)) {
        fs.rmSync(path.join(localesDir, f), { force: true })
        removed++
      }
    }
    console.log(`[afterPack] 언어팩 ${removed}개 제거 (한/영 유지)`)
  } catch (err) {
    console.log('[afterPack] 언어팩 정리 건너뜀:', err.message)
  }
}
