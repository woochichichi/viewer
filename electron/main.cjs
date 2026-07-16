// Electron 메인 프로세스
// - 파일 더블클릭 시 OS 가 넘겨주는 경로(argv / open-file / second-instance)를 받아
//   디스크에서 읽어 렌더러(React 뷰어)로 전달한다.
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('node:path')
const fs = require('node:fs')

let mainWindow = null
let rendererReady = false
let pendingPaths = [] // 렌더러 준비 전에 들어온 파일 경로 대기열

// argv 에서 실제로 존재하는 .docx/.xlsx 경로만 골라낸다.
// (exe 경로, 앱 디렉터리('.'), --flag 등은 확장자/존재검사로 자연히 걸러짐)
function pickDocPaths(argv) {
  return (argv || []).filter(
    (a) =>
      typeof a === 'string' &&
      /\.(docx|xlsx|xls|csv)$/i.test(a) &&
      fs.existsSync(a) &&
      fs.statSync(a).isFile()
  )
}

function enqueue(paths) {
  for (const p of paths) pendingPaths.push(p)
  flush()
}

function flush() {
  if (!rendererReady || !mainWindow || pendingPaths.length === 0) return
  const paths = pendingPaths
  pendingPaths = []
  for (const p of paths) {
    try {
      const buf = fs.readFileSync(p)
      // Buffer → 정확한 크기의 Uint8Array 로 변환해 IPC 전송
      const bytes = new Uint8Array(buf.byteLength)
      bytes.set(buf)
      mainWindow.webContents.send('open-file', {
        name: path.basename(p),
        data: bytes,
      })
    } catch (err) {
      mainWindow.webContents.send('open-file-error', {
        name: path.basename(p),
        message: String(err && err.message ? err.message : err),
      })
    }
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    title: '문서뷰어',
    autoHideMenuBar: true,
    backgroundColor: '#f4f5f7', // 흰 화면 깜빡임 방지
    show: false, // 렌더 준비되면 표시 → 빈 창 깜빡임 없이 매끄럽게
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })
  mainWindow.once('ready-to-show', () => mainWindow.show())
  mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  mainWindow.on('closed', () => {
    mainWindow = null
    rendererReady = false
  })
}

// 단일 인스턴스: 앱이 이미 켜져 있을 때 파일을 또 더블클릭하면
// 새 창을 띄우지 않고 기존 창에서 열도록 한다.
const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', (_event, argv) => {
    enqueue(pickDocPaths(argv))
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore()
      mainWindow.focus()
    }
  })

  // macOS: 파일 연결로 열릴 때
  app.on('open-file', (event, p) => {
    event.preventDefault()
    enqueue(pickDocPaths([p]))
  })

  app.whenReady().then(() => {
    createWindow()
    // 최초 실행 시 넘어온 파일(더블클릭으로 실행된 경우)
    enqueue(pickDocPaths(process.argv))
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })

  // 렌더러가 준비되면 대기열을 흘려보낸다.
  ipcMain.on('renderer-ready', () => {
    rendererReady = true
    flush()
  })

  // 편집 결과 저장: 저장 위치를 물어보고 파일로 기록한다.
  ipcMain.handle('save-file', async (_e, { defaultName, data }) => {
    try {
      const ext = path.extname(defaultName || '').replace('.', '') || 'xlsx'
      const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
        title: '저장',
        defaultPath: defaultName,
        filters: [{ name: ext.toUpperCase(), extensions: [ext] }],
      })
      if (canceled || !filePath) return { saved: false }
      const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
      fs.writeFileSync(filePath, Buffer.from(bytes))
      return { saved: true, path: filePath }
    } catch (err) {
      return { saved: false, error: String(err && err.message ? err.message : err) }
    }
  })

  // 바탕화면에 바로가기 만들기
  ipcMain.handle('create-desktop-shortcut', () => {
    try {
      if (process.platform !== 'win32') return { ok: false, reason: 'win-only' }
      const target = app.getPath('exe')
      const lnk = path.join(app.getPath('desktop'), '문서뷰어.lnk')
      const ok = shell.writeShortcutLink(lnk, 'create', {
        target,
        icon: target,
        iconIndex: 0,
        description: '로컬 문서 뷰어',
      })
      return { ok }
    } catch (err) {
      return { ok: false, reason: String(err && err.message ? err.message : err) }
    }
  })

  // 윈도우 기본 앱 설정 화면 열기 (기본 프로그램 지정 안내)
  ipcMain.handle('open-default-apps', async () => {
    try {
      await shell.openExternal('ms-settings:defaultapps')
      return { ok: true }
    } catch (err) {
      return { ok: false, reason: String(err && err.message ? err.message : err) }
    }
  })

  // 편집한 HTML → native docx 로 변환해 저장 (Node 환경에서 html-to-docx 실행)
  ipcMain.handle('export-docx', async (_e, { html, defaultName }) => {
    try {
      const HTMLtoDOCX = require('html-to-docx')
      const buf = await HTMLtoDOCX(html, null, {
        table: { row: { cantSplit: true } },
      })
      const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
        title: '저장',
        defaultPath: defaultName,
        filters: [{ name: 'Word 문서', extensions: ['docx'] }],
      })
      if (canceled || !filePath) return { saved: false }
      fs.writeFileSync(filePath, Buffer.from(buf))
      return { saved: true, path: filePath }
    } catch (err) {
      return { saved: false, error: String(err && err.message ? err.message : err) }
    }
  })
}
