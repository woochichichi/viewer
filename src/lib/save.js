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

// 편집한 HTML 을 .docx 로 저장.
// - Electron: 메인 프로세스에서 html-to-docx 로 native docx 생성(모든 뷰어 호환)
// - 브라우저: html-docx-js 로 폴백(워드에서 열림)
export async function exportDocxSave(defaultName, bodyHtml) {
  const full =
    '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' +
    bodyHtml +
    '</body></html>'
  const api = window.electronAPI
  if (api && api.exportDocx) {
    return (await api.exportDocx(full, defaultName)) || { saved: false }
  }
  // 브라우저에서는 native docx 생성이 불가 → 앱(.exe) 안내
  return {
    saved: false,
    appOnly: true,
    error:
      '워드(.docx) 저장은 데스크톱 앱(문서뷰어.exe)에서 지원됩니다.\n' +
      '엑셀(.xlsx) 저장은 웹에서도 가능합니다.',
  }
}
