// 렌더러(React)와 메인 프로세스 사이의 안전한 다리.
// contextIsolation 환경에서 window.electronAPI 로 제한된 기능만 노출한다.
const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronAPI', {
  isElectron: true,
  // 메인이 보내는 파일 열기 이벤트 구독
  onOpenFile: (cb) =>
    ipcRenderer.on('open-file', (_e, payload) => cb(payload)),
  onOpenFileError: (cb) =>
    ipcRenderer.on('open-file-error', (_e, payload) => cb(payload)),
  // 렌더러 준비 완료 알림 → 메인이 대기 중인 파일 전송
  notifyReady: () => ipcRenderer.send('renderer-ready'),
  // 편집 결과 저장: 저장 대화상자를 띄우고 파일로 기록
  saveFile: (defaultName, data) =>
    ipcRenderer.invoke('save-file', { defaultName, data }),
  // 편집한 HTML → native docx 저장 (메인 프로세스에서 변환)
  exportDocx: (html, defaultName) =>
    ipcRenderer.invoke('export-docx', { html, defaultName }),
})
