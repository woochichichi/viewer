import React, {
  Suspense,
  lazy,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react'

// 필요한 뷰어만 지연 로딩 → 초기 로딩/실행 속도 개선
const DocxView = lazy(() => import('./views/DocxView.jsx'))
const XlsxView = lazy(() => import('./views/XlsxView.jsx'))
const TextView = lazy(() => import('./views/TextView.jsx'))

const TEXT_EXTS = ['txt', 'log', 'md', 'ini', 'bat']
const ACCEPT = ['.docx', '.xlsx', '.xls', '.csv', ...TEXT_EXTS.map((e) => '.' + e)]

function extOf(name) {
  const i = name.lastIndexOf('.')
  return i === -1 ? '' : name.slice(i).toLowerCase()
}

function kindOf(name) {
  const e = extOf(name).slice(1)
  if (e === 'docx') return 'docx'
  if (e === 'xlsx') return 'xlsx'
  if (e === 'xls') return 'xls'
  if (e === 'csv') return 'csv'
  if (TEXT_EXTS.includes(e)) return e
  return null
}

// 스프레드시트 계열(엑셀 뷰어로 렌더)
const isSheet = (kind) => kind === 'xlsx' || kind === 'xls' || kind === 'csv'
const isText = (kind) => TEXT_EXTS.includes(kind)

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}

let idSeq = 0

export default function App() {
  const [files, setFiles] = useState([]) // { id, name, kind, size, buffer }
  const [activeId, setActiveId] = useState(null)
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef(null)

  const addFiles = useCallback(async (fileList) => {
    const incoming = Array.from(fileList)
    const accepted = []
    const rejected = []
    for (const f of incoming) {
      const kind = kindOf(f.name)
      if (!kind) {
        rejected.push(f.name)
        continue
      }
      const buffer = await f.arrayBuffer()
      accepted.push({
        id: ++idSeq,
        name: f.name,
        kind,
        size: f.size,
        buffer,
      })
    }
    if (rejected.length) {
      alert(
        `지원하지 않는 파일은 제외했습니다 (docx·xlsx·xls·csv·txt·log·md·ini·bat):\n` +
          rejected.join('\n')
      )
    }
    if (accepted.length) {
      setFiles((prev) => [...prev, ...accepted])
      setActiveId(accepted[accepted.length - 1].id)
    }
  }, [])

  // 이미 로드된 ArrayBuffer 로 파일 하나를 추가 (Electron 파일연결 열기용)
  const addBufferFile = useCallback((name, buffer, size) => {
    const kind = kindOf(name)
    if (!kind) return
    const item = {
      id: ++idSeq,
      name,
      kind,
      size: size ?? buffer.byteLength,
      buffer,
    }
    setFiles((prev) => [...prev, item])
    setActiveId(item.id)
  }, [])

  const onDrop = useCallback(
    (e) => {
      e.preventDefault()
      setDragging(false)
      if (e.dataTransfer?.files?.length) addFiles(e.dataTransfer.files)
    },
    [addFiles]
  )

  const onPick = useCallback(
    (e) => {
      if (e.target.files?.length) addFiles(e.target.files)
      e.target.value = '' // 같은 파일 재선택 허용
    },
    [addFiles]
  )

  const closeFile = useCallback(
    (id, e) => {
      e?.stopPropagation()
      setFiles((prev) => {
        const next = prev.filter((f) => f.id !== id)
        setActiveId((cur) => {
          if (cur !== id) return cur
          if (next.length === 0) return null
          const idx = prev.findIndex((f) => f.id === id)
          const fallback = next[Math.min(idx, next.length - 1)]
          return fallback ? fallback.id : next[0].id
        })
        return next
      })
    },
    []
  )

  // Electron 데스크톱 앱: 파일 더블클릭(파일연결)으로 열리면 메인 프로세스가
  // 파일을 읽어 여기로 밀어준다. 브라우저에서는 electronAPI 가 없어 무시된다.
  useEffect(() => {
    const api = window.electronAPI
    if (!api) return
    api.onOpenFile(({ name, data }) => {
      const u8 = data instanceof Uint8Array ? data : new Uint8Array(data)
      // 정확한 크기의 독립 ArrayBuffer 확보
      const buffer = u8.slice().buffer
      addBufferFile(name, buffer)
    })
    api.onOpenFileError?.(({ name, message }) => {
      alert(`파일을 열 수 없습니다: ${name}\n${message}`)
    })
    api.notifyReady()
  }, [addBufferFile])

  // ---- 확대/축소 ----
  // 파일별 확대 배율 (탭마다 개별 기억)
  const [zoomMap, setZoomMap] = useState({})
  const clampZoom = (z) => Math.min(3, Math.max(0.3, Math.round(z * 100) / 100))
  const setActiveZoom = useCallback(
    (v) => {
      setZoomMap((m) => {
        const cur = m[activeId] ?? 1
        const next = clampZoom(typeof v === 'function' ? v(cur) : v)
        return { ...m, [activeId]: next }
      })
    },
    [activeId]
  )
  const zoomIn = useCallback(() => setActiveZoom((z) => z + 0.1), [setActiveZoom])
  const zoomOut = useCallback(() => setActiveZoom((z) => z - 0.1), [setActiveZoom])
  const zoomReset = useCallback(() => setActiveZoom(1), [setActiveZoom])
  const zoom = zoomMap[activeId] ?? 1

  // ---- 찾기(Ctrl+F) ----
  const [findOpen, setFindOpen] = useState(false)
  const [findQuery, setFindQuery] = useState('')
  const findInputRef = useRef(null)
  const runFind = useCallback((backwards) => {
    const q = findInputRef.current?.value?.trim()
    if (!q) return
    // Electron/Chromium 의 내장 찾기 사용 (해당 위치로 스크롤 + 하이라이트)
    const sel = window.getSelection?.()
    if (sel) sel.collapseToStart?.()
    window.find?.(q, false, backwards, true, false, false, false)
  }, [])

  useEffect(() => {
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault()
        setFindOpen(true)
        setTimeout(() => findInputRef.current?.select(), 30)
        return
      }
      if (e.key === 'Escape' && findOpen) {
        setFindOpen(false)
        return
      }
      if (!(e.ctrlKey || e.metaKey)) return
      if (e.key === '=' || e.key === '+') {
        e.preventDefault()
        zoomIn()
      } else if (e.key === '-') {
        e.preventDefault()
        zoomOut()
      } else if (e.key === '0') {
        e.preventDefault()
        zoomReset()
      }
    }
    const onWheel = (e) => {
      if (!(e.ctrlKey || e.metaKey)) return
      e.preventDefault()
      setActiveZoom((z) => z + (e.deltaY < 0 ? 0.1 : -0.1))
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('wheel', onWheel, { passive: false })
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('wheel', onWheel)
    }
  }, [zoomIn, zoomOut, zoomReset, findOpen])

  // ---- 우클릭 컨텍스트 메뉴 ----
  const [ctx, setCtx] = useState(null) // { x, y, text }
  const onContextMenu = useCallback((e) => {
    const sel = window.getSelection?.()
    const text = sel && !sel.isCollapsed ? sel.toString().trim() : ''
    e.preventDefault()
    const x = Math.min(e.clientX, window.innerWidth - 210)
    const y = Math.min(e.clientY, window.innerHeight - 160)
    setCtx({ x, y, text })
  }, [])
  useEffect(() => {
    if (!ctx) return
    const close = () => setCtx(null)
    const onEsc = (e) => e.key === 'Escape' && setCtx(null)
    window.addEventListener('click', close)
    window.addEventListener('scroll', close, true)
    window.addEventListener('keydown', onEsc)
    return () => {
      window.removeEventListener('click', close)
      window.removeEventListener('scroll', close, true)
      window.removeEventListener('keydown', onEsc)
    }
  }, [ctx])
  const ctxCopy = () => {
    document.execCommand('copy') // 선택 영역 복사 (docx 개행 보정 핸들러도 적용됨)
    setCtx(null)
  }
  const ctxSelectAll = () => {
    const el =
      document.querySelector('.render-area .docx-host') ||
      document.querySelector('.render-area .docx-edit') ||
      document.querySelector('.render-area .xlsx-table')
    if (el) {
      const r = document.createRange()
      r.selectNodeContents(el)
      const s = window.getSelection()
      s.removeAllRanges()
      s.addRange(r)
    }
    setCtx(null)
  }
  const ctxFind = () => {
    const t = ctx?.text || ''
    setCtx(null)
    setFindQuery(t)
    setFindOpen(true)
    setTimeout(() => {
      if (findInputRef.current) {
        findInputRef.current.value = t
        findInputRef.current.focus()
      }
      runFind(false)
    }, 40)
  }

  // ---- 최초 실행 안내(데스크톱 앱, 윈도우) ----
  const [welcome, setWelcome] = useState(false)
  useEffect(() => {
    const api = window.electronAPI
    if (api && api.platform === 'win32' && !localStorage.getItem('dv-welcomed')) {
      setWelcome(true)
    }
  }, [])
  const dismissWelcome = () => {
    try {
      localStorage.setItem('dv-welcomed', '1')
    } catch {}
    setWelcome(false)
  }
  const makeShortcut = async () => {
    const r = await window.electronAPI?.createDesktopShortcut?.()
    alert(
      r?.ok
        ? '바탕화면에 "문서뷰어" 바로가기를 만들었습니다.\n아이콘 우클릭 → "작업 표시줄에 고정"으로 작업표시줄에도 넣을 수 있어요.'
        : '바로가기 생성에 실패했습니다: ' + (r?.reason || '알 수 없음')
    )
  }
  const openDefaults = async () => {
    await window.electronAPI?.openDefaultApps?.()
    alert(
      '윈도우 "기본 앱" 설정이 열립니다.\n' +
        '.docx / .xlsx 항목에서 "문서뷰어"를 선택하면 기본 프로그램이 됩니다.'
    )
  }

  const active = files.find((f) => f.id === activeId) || null

  return (
    <div
      className="app"
      onDragOver={(e) => {
        e.preventDefault()
        if (!dragging) setDragging(true)
      }}
      onDragLeave={(e) => {
        // 실제로 창을 벗어날 때만 해제
        if (e.currentTarget === e.target) setDragging(false)
      }}
      onDrop={onDrop}
    >
      <aside className="sidebar">
        <div className="sidebar-head">
          <h1>문서 뷰어</h1>
          <p className="hint">
            v{__APP_VERSION__} · 문서·표·텍스트 · 오프라인
          </p>
        </div>

        <button className="pick-btn" onClick={() => inputRef.current?.click()}>
          + 파일 열기
        </button>
        <input
          ref={inputRef}
          type="file"
          accept={ACCEPT.join(',')}
          multiple
          hidden
          onChange={onPick}
        />

        <ul className="file-list">
          {files.length === 0 && (
            <li className="empty-note">열린 파일이 없습니다</li>
          )}
          {files.map((f) => (
            <li
              key={f.id}
              className={
                'file-item' + (f.id === activeId ? ' active' : '')
              }
              onClick={() => setActiveId(f.id)}
              title={f.name}
            >
              <span className={'badge badge-' + f.kind}>{f.kind}</span>
              <span className="file-name">{f.name}</span>
              <span className="file-size">{formatSize(f.size)}</span>
              <button
                className="close-x"
                onClick={(e) => closeFile(f.id, e)}
                title="닫기"
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      </aside>

      <main className="viewer">
        {/* 상단 탭 */}
        {files.length > 0 && (
          <div className="tabbar">
            {files.map((f) => (
              <div
                key={f.id}
                className={'tab' + (f.id === activeId ? ' active' : '')}
                onClick={() => setActiveId(f.id)}
                title={f.name}
              >
                <span className={'badge badge-' + f.kind}>{f.kind}</span>
                <span className="tab-name">{f.name}</span>
                <button
                  className="close-x"
                  onClick={(e) => closeFile(f.id, e)}
                  title="닫기"
                >
                  ×
                </button>
              </div>
            ))}
          </div>
        )}

        <div className="render-area" onContextMenu={onContextMenu}>
          {active && findOpen && (
            <div className="find-bar" onKeyDown={(e) => e.stopPropagation()}>
              <input
                ref={findInputRef}
                className="find-input"
                placeholder="찾기…"
                value={findQuery}
                onChange={(e) => setFindQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') runFind(e.shiftKey)
                  if (e.key === 'Escape') setFindOpen(false)
                }}
              />
              <button onClick={() => runFind(true)} title="이전">
                ↑
              </button>
              <button onClick={() => runFind(false)} title="다음 (Enter)">
                ↓
              </button>
              <button
                className="find-close"
                onClick={() => setFindOpen(false)}
                title="닫기 (Esc)"
              >
                ×
              </button>
            </div>
          )}
          {!active && (
            <div className="dropzone">
              <div className="dropzone-inner">
                <div className="dz-icon">📄</div>
                <p className="dz-title">
                  파일을 이곳에 끌어다 놓으세요
                </p>
                <p className="dz-sub">
                  또는 좌측의 <b>파일 열기</b> 버튼을 사용하세요
                  <br />
                  지원 형식: docx · xlsx · xls · csv · txt · log · md · ini · bat
                </p>
                <div className="dz-tips">
                  <span>🔍 찾기 <kbd>Ctrl</kbd>+<kbd>F</kbd></span>
                  <span>🔎 확대/축소 <kbd>Ctrl</kbd>+<kbd>+</kbd>/<kbd>−</kbd></span>
                  <span>✏️ 열고 나서 편집·저장</span>
                </div>
              </div>
            </div>
          )}
          {active && (
            <Suspense
              fallback={<div className="state-note">여는 중…</div>}
            >
              {active.kind === 'docx' && (
                <DocxView
                  key={active.id}
                  buffer={active.buffer}
                  name={active.name}
                  zoom={zoom}
                  onAutoFit={setActiveZoom}
                />
              )}
              {isSheet(active.kind) && (
                <XlsxView
                  key={active.id}
                  buffer={active.buffer}
                  name={active.name}
                  kind={active.kind}
                  zoom={zoom}
                />
              )}
              {isText(active.kind) && (
                <TextView
                  key={active.id}
                  buffer={active.buffer}
                  name={active.name}
                  kind={active.kind}
                  zoom={zoom}
                />
              )}
            </Suspense>
          )}

          {active && (
            <div className="zoom-bar">
              <button onClick={zoomOut} title="축소 (Ctrl -)">
                −
              </button>
              <button
                className="zoom-pct"
                onClick={zoomReset}
                title="원래대로 (Ctrl 0)"
              >
                {Math.round(zoom * 100)}%
              </button>
              <button onClick={zoomIn} title="확대 (Ctrl +)">
                +
              </button>
            </div>
          )}
        </div>
      </main>

      {dragging && (
        <div className="drag-overlay">
          <div className="drag-overlay-msg">여기에 놓으세요</div>
        </div>
      )}

      {ctx && (
        <div
          className="ctx-menu"
          style={{ left: ctx.x, top: ctx.y }}
          onClick={(e) => e.stopPropagation()}
          onContextMenu={(e) => e.preventDefault()}
        >
          <button className="ctx-item" disabled={!ctx.text} onClick={ctxCopy}>
            <span>복사</span>
            <span className="ctx-key">Ctrl+C</span>
          </button>
          <button className="ctx-item" disabled={!ctx.text} onClick={ctxFind}>
            <span>
              {ctx.text
                ? `"${ctx.text.slice(0, 12)}${ctx.text.length > 12 ? '…' : ''}" 찾기`
                : '찾기'}
            </span>
            <span className="ctx-key">Ctrl+F</span>
          </button>
          <div className="ctx-sep" />
          <button className="ctx-item" onClick={ctxSelectAll}>
            <span>전체 선택</span>
            <span className="ctx-key">Ctrl+A</span>
          </button>
        </div>
      )}

      {welcome && (
        <div className="welcome-overlay">
          <div className="welcome-card">
            <h2>문서뷰어 설치를 완료했어요 🎉</h2>
            <p className="welcome-sub">
              아래 설정을 해두면 훨씬 편하게 쓸 수 있어요. (선택)
            </p>
            <div className="welcome-actions">
              <button className="wc-btn primary" onClick={makeShortcut}>
                🖥️ 바탕화면 바로가기 만들기 <span className="rec">권장</span>
              </button>
              <button className="wc-btn" onClick={openDefaults}>
                📎 docx·xlsx 기본 프로그램으로 설정 <span className="rec">권장</span>
              </button>
            </div>
            <p className="welcome-tip">
              작업 표시줄에 고정: 바탕화면 아이콘 <b>우클릭 → "작업 표시줄에 고정"</b>
            </p>
            <button className="wc-close" onClick={dismissWelcome}>
              닫기
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
