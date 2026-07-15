import React, { useCallback, useRef, useState } from 'react'
import DocxView from './views/DocxView.jsx'
import XlsxView from './views/XlsxView.jsx'

const ACCEPT = ['.docx', '.xlsx']

function extOf(name) {
  const i = name.lastIndexOf('.')
  return i === -1 ? '' : name.slice(i).toLowerCase()
}

function kindOf(name) {
  const e = extOf(name)
  if (e === '.docx') return 'docx'
  if (e === '.xlsx') return 'xlsx'
  return null
}

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
        `지원하지 않는 파일은 제외했습니다 (.docx / .xlsx 만 가능):\n` +
          rejected.join('\n')
      )
    }
    if (accepted.length) {
      setFiles((prev) => [...prev, ...accepted])
      setActiveId(accepted[accepted.length - 1].id)
    }
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
          <p className="hint">.docx · .xlsx · 오프라인</p>
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

        <div className="render-area">
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
                  지원 형식: .docx, .xlsx
                </p>
              </div>
            </div>
          )}
          {active && active.kind === 'docx' && (
            <DocxView key={active.id} buffer={active.buffer} />
          )}
          {active && active.kind === 'xlsx' && (
            <XlsxView key={active.id} buffer={active.buffer} />
          )}
        </div>
      </main>

      {dragging && (
        <div className="drag-overlay">
          <div className="drag-overlay-msg">여기에 놓으세요</div>
        </div>
      )}
    </div>
  )
}
