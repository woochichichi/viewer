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

// 두 ArrayBuffer 내용이 같은지 (같은 파일 중복 열기 판별)
function sameBytes(a, b) {
  if (a === b) return true
  if (!a || !b || a.byteLength !== b.byteLength) return false
  const x = new Uint8Array(a)
  const y = new Uint8Array(b)
  for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return false
  return true
}

// 파일 끌어오기인지(텍스트·이미지 드래그와 구분)
const hasFiles = (e) => Array.from(e.dataTransfer?.types || []).includes('Files')
const isEditable = (t) =>
  !!t && (t.isContentEditable || t.tagName === 'TEXTAREA' || t.tagName === 'INPUT')

// 처음 열 때 확대 배율 (형식별). 워드는 1920 모니터에서도 A4 가 과하게 크지 않게 90%
const DEFAULT_ZOOM = { docx: 0.9 }

export default function App() {
  const [files, setFiles] = useState([]) // { id, name, kind, size, buffer }
  const [activeId, setActiveId] = useState(null)
  const mountedRef = useRef(new Set()) // 뷰가 마운트된 파일 id
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef(null)

  // 파일별 확대 배율 (탭마다 개별 기억). 워드는 열 때 기본 90% (DEFAULT_ZOOM)
  const [zoomMap, setZoomMap] = useState({})

  // 현재 열린 파일 목록(중복 열기 검사용 — setState 비동기 대비 ref 로 동기화)
  const filesRef = useRef([])
  useEffect(() => {
    filesRef.current = files
  }, [files])

  // 열린 파일 추가. 이름·내용이 같은 파일이 이미 열려 있으면 새 탭 대신 그 탭으로 전환
  // (워드/엑셀처럼 같은 문서를 두 번 열지 않음)
  const pushItems = useCallback((items) => {
    const opened = [...filesRef.current]
    const fresh = []
    let lastId = null
    for (const it of items) {
      const dup = opened.find(
        (f) => f.name === it.name && sameBytes(f.buffer, it.buffer)
      )
      if (dup) {
        lastId = dup.id
        continue
      }
      opened.push(it)
      fresh.push(it)
      lastId = it.id
    }
    filesRef.current = opened
    if (fresh.length) {
      setFiles((prev) => [...prev, ...fresh])
      const init = fresh.filter((f) => DEFAULT_ZOOM[f.kind])
      if (init.length)
        setZoomMap((m) => {
          const next = { ...m }
          for (const f of init) next[f.id] = DEFAULT_ZOOM[f.kind]
          return next
        })
    }
    if (lastId != null) setActiveId(lastId)
  }, [])

  const addFiles = useCallback(
    async (fileList) => {
      const incoming = Array.from(fileList)
      const accepted = []
      const rejected = []
      const failed = []
      for (const f of incoming) {
        const kind = kindOf(f.name)
        if (!kind) {
          rejected.push(f.name)
          continue
        }
        try {
          const buffer = await f.arrayBuffer()
          accepted.push({
            id: ++idSeq,
            name: f.name,
            kind,
            size: f.size,
            buffer,
          })
        } catch (err) {
          // 폴더·삭제된 파일·권한 없음 등 → 그 파일만 건너뛰고 나머지는 연다
          failed.push(`${f.name} (${err?.message || err})`)
        }
      }
      if (rejected.length) {
        alert(
          `지원하지 않는 파일은 제외했습니다 (docx·xlsx·xls·csv·txt·log·md·ini·bat):\n` +
            rejected.join('\n')
        )
      }
      if (failed.length) {
        alert('파일을 읽을 수 없습니다:\n' + failed.join('\n'))
      }
      if (accepted.length) pushItems(accepted)
    },
    [pushItems]
  )

  // 이미 로드된 ArrayBuffer 로 파일 하나를 추가 (Electron 파일연결 열기용)
  const addBufferFile = useCallback(
    (name, buffer, size) => {
      const kind = kindOf(name)
      if (!kind) return
      pushItems([
        {
          id: ++idSeq,
          name,
          kind,
          size: size ?? buffer.byteLength,
          buffer,
        },
      ])
    },
    [pushItems]
  )

  // dragenter/leave 는 자식 요소마다 발생 → 깊이 카운트로 창을 실제로 벗어났는지 판단
  const dragDepth = useRef(0)
  const onDrop = useCallback(
    (e) => {
      dragDepth.current = 0
      setDragging(false)
      // 파일이 아닌 드롭: 편집 영역 안(텍스트 끌어 옮기기)은 기본 동작 유지,
      // 그 밖(링크·텍스트 드롭)은 막아 앱 화면이 다른 주소로 넘어가지 않게
      if (!hasFiles(e)) {
        if (!isEditable(e.target)) e.preventDefault()
        return
      }
      e.preventDefault()
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

  // 탭·목록 항목 키보드 조작: Enter/Space=선택, Delete=닫기
  const onItemKey = (e, id) => {
    if (e.target !== e.currentTarget) return
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault()
      setActiveId(id)
    } else if (e.key === 'Delete') {
      e.preventDefault()
      closeFile(id)
    }
  }

  const closeFile = useCallback(
    (id, e) => {
      e?.stopPropagation()
      mountedRef.current.delete(id)
      setZoomMap((m) => {
        if (!(id in m)) return m
        const { [id]: _, ...rest } = m
        return rest
      })
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
    const off1 = api.onOpenFile(({ name, data }) => {
      const u8 = data instanceof Uint8Array ? data : new Uint8Array(data)
      // 정확한 크기의 독립 ArrayBuffer 확보
      const buffer = u8.slice().buffer
      addBufferFile(name, buffer)
    })
    const off2 = api.onOpenFileError?.(({ name, message }) => {
      alert(`파일을 열 수 없습니다: ${name}\n${message}`)
    })
    api.notifyReady()
    // 리스너 해제(재마운트 시 같은 파일이 두 번 열리는 것 방지)
    return () => {
      if (typeof off1 === 'function') off1()
      if (typeof off2 === 'function') off2()
    }
  }, [addBufferFile])

  // ---- 확대/축소 ----
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
  // 특정 파일의 배율 지정 (docx '페이지 맞춤' 버튼용 — 그 사이 다른 탭으로 바뀌어도 정확히)
  const setZoomFor = useCallback((id, z) => {
    setZoomMap((m) => ({ ...m, [id]: clampZoom(z) }))
  }, [])
  const zoomIn = useCallback(() => setActiveZoom((z) => z + 0.1), [setActiveZoom])
  const zoomOut = useCallback(() => setActiveZoom((z) => z - 0.1), [setActiveZoom])
  const zoomReset = useCallback(() => setActiveZoom(1), [setActiveZoom])
  const zoom = zoomMap[activeId] ?? 1

  // ---- 찾기(Ctrl+F) ----
  const [findOpen, setFindOpen] = useState(false)
  const [findQuery, setFindQuery] = useState('')
  const findInputRef = useRef(null)
  const [findMiss, setFindMiss] = useState(false)
  const runFind = useCallback((backwards) => {
    const q = findInputRef.current?.value?.trim()
    if (!q) return
    // Electron/Chromium 의 내장 찾기 사용 (해당 위치로 스크롤 + 하이라이트)
    // 검색 범위 = 현재 문서 영역(찾기 입력창 자체는 제외)
    const root = document.querySelector('.render-area > .view-pane.active')
    const sel = window.getSelection?.()
    if (!root || !sel || !window.find) return
    const inRoot = () => sel.rangeCount > 0 && root.contains(sel.anchorNode)
    // 문서 처음(다음 찾기) 또는 끝(이전 찾기)으로 캐럿 이동
    const toEdge = () => {
      const r = document.createRange()
      r.selectNodeContents(root)
      r.collapse(!backwards)
      sel.removeAllRanges()
      sel.addRange(r)
    }
    // 현재 선택(직전 결과) 다음부터 찾는다. 선택이 문서 밖(입력창 등)이면 처음부터.
    if (!inRoot()) toEdge()
    let ok = window.find(q, false, backwards, false, false, false, false) && inRoot()
    if (!ok) {
      // 끝까지 갔으면 처음(이전 찾기는 끝)부터 다시 — 순환 검색
      toEdge()
      ok = window.find(q, false, backwards, false, false, false, false) && inRoot()
    }
    if (!ok) sel.removeAllRanges()
    setFindMiss(!ok)
  }, [])

  // ---- 인쇄(Ctrl+P) ----
  // 화면 장식(사이드바·탭·툴바 등)은 @media print CSS 로 숨기고 문서만 100% 배율로 인쇄.
  // window.print() 는 브라우저/단일 HTML/Electron 모두 동일하게 동작(시스템 인쇄 대화상자).
  const doPrint = useCallback(() => {
    setCtx(null)
    if (!activeId) return
    window.print()
  }, [activeId])

  useEffect(() => {
    // 한글 입력 상태(ㅔ/ㄹ)에서도 동작하도록 e.code 도 함께 확인
    const isKey = (e, ch) =>
      e.key === ch || e.key === ch.toUpperCase() || e.code === 'Key' + ch.toUpperCase()
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && isKey(e, 'p')) {
        e.preventDefault() // 브라우저 기본 인쇄 대신 앱 인쇄
        doPrint()
        return
      }
      if ((e.ctrlKey || e.metaKey) && !e.altKey && isKey(e, 'f')) {
        e.preventDefault()
        setFindOpen(true)
        // 이미 열려 있으면 바로 선택, 새로 열리면 입력창 autoFocus 로 포커스
        findInputRef.current?.select()
        return
      }
      // 파일 열기 (브라우저 기본 Ctrl+O 는 앱을 떠나 파일을 열어버림)
      if ((e.ctrlKey || e.metaKey) && !e.altKey && isKey(e, 'o')) {
        e.preventDefault()
        inputRef.current?.click()
        return
      }
      // 저장: 활성 탭의 뷰에 저장 요청 (브라우저 기본 '페이지 저장' 대신)
      if ((e.ctrlKey || e.metaKey) && !e.altKey && isKey(e, 's')) {
        e.preventDefault()
        window.dispatchEvent(new CustomEvent('dv-save'))
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
    // 마우스 휠은 한 칸(≈100)마다 10%. 터치패드 핀치는 작은 값이 연속으로 오므로
    // 누적해서 일정량마다 10% (이벤트마다 10%씩이면 순식간에 300%/30% 로 튐)
    let wheelAcc = 0
    const onWheel = (e) => {
      if (!(e.ctrlKey || e.metaKey)) return
      e.preventDefault()
      wheelAcc += e.deltaMode === 1 ? e.deltaY * 33 : e.deltaY
      if (Math.abs(wheelAcc) < 50) return
      const dir = wheelAcc < 0 ? 0.1 : -0.1
      wheelAcc = 0
      setActiveZoom((z) => z + dir)
    }
    // 캡처 단계에서 받는다: 찾기 입력창 등이 전파를 막아도 Ctrl+P/F/확대 단축키가
    // 항상 앱 처리로 가도록 (막히면 브라우저 기본 인쇄가 뜨거나 Electron 에선 무반응)
    window.addEventListener('keydown', onKey, true)
    window.addEventListener('wheel', onWheel, { passive: false })
    return () => {
      window.removeEventListener('keydown', onKey, true)
      window.removeEventListener('wheel', onWheel)
    }
  }, [zoomIn, zoomOut, zoomReset, findOpen, doPrint])

  // ---- 우클릭 컨텍스트 메뉴 ----
  const [ctx, setCtx] = useState(null) // { x, y, text }
  const onContextMenu = useCallback((e) => {
    const sel = window.getSelection?.()
    const text = sel && !sel.isCollapsed ? sel.toString().trim() : ''
    e.preventDefault()
    const x = Math.max(4, Math.min(e.clientX, window.innerWidth - 210))
    const y = Math.max(4, Math.min(e.clientY, window.innerHeight - 200))
    setCtx({ x, y, text })
  }, [])
  useEffect(() => {
    if (!ctx) return
    const close = () => setCtx(null)
    // 열기 직전의 스크롤(휠 직후 우클릭, 편집 영역 포커스 이동)로 생긴 늦은 scroll 이벤트에
    // 바로 닫히지 않도록 잠깐 무시
    const openedAt = performance.now()
    const onScroll = () => performance.now() - openedAt > 150 && setCtx(null)
    // 메뉴 밖을 누르면(좌/우클릭 모두) 닫기. 우클릭은 mousedown 이 contextmenu 보다
    // 먼저 오므로 다른 곳 우클릭 시 닫힌 뒤 새 위치에 다시 열린다.
    const onDown = (e) => {
      if (!e.target.closest?.('.ctx-menu')) setCtx(null)
    }
    const onEsc = (e) => e.key === 'Escape' && setCtx(null)
    window.addEventListener('mousedown', onDown, true)
    window.addEventListener('scroll', onScroll, true)
    window.addEventListener('keydown', onEsc)
    window.addEventListener('blur', close)
    window.addEventListener('resize', close)
    return () => {
      window.removeEventListener('mousedown', onDown, true)
      window.removeEventListener('scroll', onScroll, true)
      window.removeEventListener('keydown', onEsc)
      window.removeEventListener('blur', close)
      window.removeEventListener('resize', close)
    }
  }, [ctx])
  const ctxCopy = () => {
    document.execCommand('copy') // 선택 영역 복사 (docx 개행 보정 핸들러도 적용됨)
    setCtx(null)
  }
  const ctxSelectAll = () => {
    // 현재 화면에 보이는 문서 영역만 (docx 보기/편집 영역은 둘 다 마운트돼 있음)
    const el = [
      ...document.querySelectorAll(
        '.render-area .docx-host, .render-area .docx-edit, .render-area .xlsx-table,' +
          ' .render-area .md-view, .render-area .text-view:not(.print-only), .render-area .text-edit'
      ),
    ].find((n) => n.getClientRects().length > 0)
    if (el && el.tagName === 'TEXTAREA') {
      el.focus()
      el.select()
    } else if (el) {
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
    let welcomed = false
    try {
      welcomed = !!localStorage.getItem('dv-welcomed')
    } catch {}
    if (api && api.platform === 'win32' && !welcomed) setWelcome(true)
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

  // 문서 안 링크 클릭: 앱 화면이 링크 주소로 넘어가 열린 파일이 모두 사라지는 것 방지.
  // 문서 내부 책갈피(#...)만 그대로 두고, 외부 링크는 새 창(Electron 은 기본 브라우저)으로.
  const onDocClick = useCallback((e) => {
    const a = e.target.closest?.('a[href]')
    if (!a || a.isContentEditable) return
    const href = a.getAttribute('href') || ''
    if (href.startsWith('#')) return
    e.preventDefault()
    if (/^(https?|mailto):/i.test(href.trim())) {
      window.open(href.trim(), '_blank', 'noopener,noreferrer')
    }
  }, [])

  const active = files.find((f) => f.id === activeId) || null

  // 창 제목 = 현재 파일 이름 (작업 표시줄·브라우저 탭에서 구분)
  useEffect(() => {
    document.title = active ? `${active.name} - 문서뷰어` : '문서뷰어'
  }, [active])
  // 한 번이라도 활성화된 파일만 뷰를 마운트(여러 파일을 한꺼번에 열어도 보는 것만 파싱)
  if (active) mountedRef.current.add(active.id)

  return (
    <div
      className="app"
      onDragEnter={(e) => {
        if (!hasFiles(e)) return
        dragDepth.current++
        if (!dragging) setDragging(true)
      }}
      onDragOver={(e) => {
        // 기본 동작(파일을 창에서 열어 앱이 사라짐) 방지
        if (hasFiles(e) || !isEditable(e.target)) e.preventDefault()
      }}
      onDragLeave={(e) => {
        if (!hasFiles(e)) return
        // 실제로 창을 벗어날 때만 해제 (자식 요소 간 이동은 무시)
        dragDepth.current = Math.max(0, dragDepth.current - 1)
        if (dragDepth.current === 0) setDragging(false)
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

        <button
          className="pick-btn"
          onClick={() => inputRef.current?.click()}
          title="파일 열기 (Ctrl+O)"
        >
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
              onKeyDown={(e) => onItemKey(e, f.id)}
              tabIndex={0}
              title={`${f.name} (${formatSize(f.size)})`}
            >
              <span className={'badge badge-' + f.kind}>{f.kind}</span>
              <span className="file-name">{f.name}</span>
              <span className="file-size">{formatSize(f.size)}</span>
              <button
                className="close-x"
                onClick={(e) => closeFile(f.id, e)}
                title="닫기"
                aria-label={`${f.name} 닫기`}
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
          <div className="tabbar" role="tablist">
            {files.map((f) => (
              <div
                key={f.id}
                className={'tab' + (f.id === activeId ? ' active' : '')}
                onClick={() => setActiveId(f.id)}
                onKeyDown={(e) => onItemKey(e, f.id)}
                // 가운데 버튼 클릭 = 탭 닫기 (브라우저 탭과 동일)
                onMouseDown={(e) => e.button === 1 && e.preventDefault()}
                onAuxClick={(e) => e.button === 1 && closeFile(f.id, e)}
                role="tab"
                aria-selected={f.id === activeId}
                tabIndex={0}
                title={f.name}
              >
                <span className={'badge badge-' + f.kind}>{f.kind}</span>
                <span className="tab-name">{f.name}</span>
                <button
                  className="close-x"
                  onClick={(e) => closeFile(f.id, e)}
                  title="닫기 (가운데 버튼 클릭)"
                  aria-label={`${f.name} 닫기`}
                >
                  ×
                </button>
              </div>
            ))}
          </div>
        )}

        <div
          className="render-area"
          onContextMenu={onContextMenu}
          onClickCapture={onDocClick}
        >
          {active && findOpen && (
            <div className="find-bar" onKeyDown={(e) => e.stopPropagation()}>
              <input
                ref={findInputRef}
                className={'find-input' + (findMiss ? ' miss' : '')}
                placeholder="문서에서 찾기…"
                autoFocus
                aria-label="찾을 내용"
                value={findQuery}
                onChange={(e) => {
                  setFindQuery(e.target.value)
                  setFindMiss(false)
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') runFind(e.shiftKey)
                  if (e.key === 'Escape') setFindOpen(false)
                }}
              />
              {findMiss && <span className="find-miss">결과 없음</span>}
              <button onClick={() => runFind(true)} title="이전 (Shift+Enter)">
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
                <p className="dz-sub">여러 파일을 한 번에 열 수 있어요</p>
                <button
                  className="dz-open"
                  onClick={() => inputRef.current?.click()}
                  title="파일 열기 (Ctrl+O)"
                >
                  파일 열기…
                </button>
                <div className="dz-formats">
                  <span className="badge badge-docx">docx</span>
                  <span className="badge badge-xlsx">xlsx</span>
                  <span className="badge badge-xls">xls</span>
                  <span className="badge badge-csv">csv</span>
                  {TEXT_EXTS.map((e) => (
                    <span key={e} className={'badge badge-' + e}>
                      {e}
                    </span>
                  ))}
                </div>
                <div className="dz-tips">
                  <span>파일 열기</span>
                  <span>
                    <kbd>Ctrl</kbd>+<kbd>O</kbd>
                  </span>
                  <span>찾기</span>
                  <span>
                    <kbd>Ctrl</kbd>+<kbd>F</kbd>
                  </span>
                  <span>확대 / 축소</span>
                  <span>
                    <kbd>Ctrl</kbd>+<kbd>+</kbd> / <kbd>−</kbd> · <kbd>Ctrl</kbd>+휠
                  </span>
                  <span>저장 (편집 후)</span>
                  <span>
                    <kbd>Ctrl</kbd>+<kbd>S</kbd>
                  </span>
                  <span>인쇄</span>
                  <span>
                    <kbd>Ctrl</kbd>+<kbd>P</kbd>
                  </span>
                </div>
                <p className="dz-note">
                  파일은 이 PC 안에서만 열리며 외부로 전송되지 않습니다.
                </p>
              </div>
            </div>
          )}
          {/* 한 번 연 탭은 계속 마운트해 두고 숨김만 전환 → 탭을 오가도 편집 내용·편집 모드·
              스크롤·확대 배율이 유지된다(전에는 탭 전환 시 저장 안 한 편집이 사라졌음) */}
          {files
            .filter((f) => mountedRef.current.has(f.id))
            .map((f) => (
              <div
                key={f.id}
                className={'view-pane' + (f.id === activeId ? ' active' : '')}
              >
                <Suspense
                  fallback={<div className="state-note loading">여는 중…</div>}
                >
                  {f.kind === 'docx' && (
                    <DocxView
                      buffer={f.buffer}
                      name={f.name}
                      docId={f.id}
                      zoom={zoomMap[f.id] ?? 1}
                      onAutoFit={(z) => setZoomFor(f.id, z)}
                      onPrint={doPrint}
                    />
                  )}
                  {isSheet(f.kind) && (
                    <XlsxView
                      buffer={f.buffer}
                      name={f.name}
                      kind={f.kind}
                      zoom={zoomMap[f.id] ?? 1}
                      onPrint={doPrint}
                    />
                  )}
                  {isText(f.kind) && (
                    <TextView
                      buffer={f.buffer}
                      name={f.name}
                      kind={f.kind}
                      zoom={zoomMap[f.id] ?? 1}
                      onPrint={doPrint}
                    />
                  )}
                </Suspense>
              </div>
            ))}

          {active && (
            <div className="zoom-bar">
              <button onClick={zoomOut} title="축소 (Ctrl+−, Ctrl+휠)">
                −
              </button>
              <button
                className="zoom-pct"
                onClick={zoomReset}
                title="100%로 (Ctrl+0)"
              >
                {Math.round(zoom * 100)}%
              </button>
              <button onClick={zoomIn} title="확대 (Ctrl++, Ctrl+휠)">
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
          <div className="ctx-sep" />
          <button className="ctx-item" disabled={!active} onClick={doPrint}>
            <span>🖨 인쇄</span>
            <span className="ctx-key">Ctrl+P</span>
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
