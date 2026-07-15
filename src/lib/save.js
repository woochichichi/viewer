// 편집 결과를 파일로 저장.
// - Electron: 저장 대화상자(메인 프로세스) 사용
// - 브라우저: Blob 다운로드로 폴백
export async function saveBytes(defaultName, bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const api = window.electronAPI
  if (api && api.saveFile) {
    const res = await api.saveFile(defaultName, u8)
    return res || { saved: false }
  }
  // 브라우저 폴백
  const blob = new Blob([u8], {
    type: 'application/octet-stream',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = defaultName
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
  return { saved: true, path: defaultName }
}

// 원본 파일명에 접미사 붙이기: report.xlsx → report-편집.xlsx
export function editedName(name, suffix = '-편집') {
  const i = name.lastIndexOf('.')
  if (i === -1) return name + suffix
  return name.slice(0, i) + suffix + name.slice(i)
}
