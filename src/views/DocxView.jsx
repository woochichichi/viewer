import React, { useCallback, useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { exportDocxSave, editedName } from '../lib/save.js'

// 보기: docx-preview(서식 충실). 편집: mammoth 로 단순 HTML 화 → 편집 → docx 저장.
export default function DocxView({ buffer, name = 'document.docx', zoom = 1 }) {
  const containerRef = useRef(null) // docx-preview 보기 영역
  const editRef = useRef(null) // 편집(contentEditable) 영역
  const [status, setStatus] = useState('loading') // loading | done | error
  const [error, setError] = useState('')
  const [editMode, setEditMode] = useState(false)
  const [editHtml, setEditHtml] = useState(null) // mammoth 결과(최초 1회)
  const [preparing, setPreparing] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)

  // ---- 보기: docx-preview 렌더 ----
  useEffect(() => {
    let cancelled = false
    const el = containerRef.current
    if (!el) return
    el.innerHTML = ''
    setStatus('loading')
    setError('')
    renderAsync(buffer, el, undefined, {
      className: 'docx',
      inWrapper: true,
      breakPages: true,
      experimental: true,
      useBase64URL: true,
    })
      .then(() => !cancelled && setStatus('done'))
      .catch((err) => {
        if (cancelled) return
        console.error(err)
        setError(err?.message || String(err))
        setStatus('error')
      })
    return () => {
      cancelled = true
    }
  }, [buffer])

  // ---- 복사 시 문단마다 빈 줄 끼는 문제 보정 (보기 영역) ----
  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const onCopy = (e) => {
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return
      if (!el.contains(sel.anchorNode) && !el.contains(sel.focusNode)) return
      const text = sel
        .toString()
        .replace(/[ \t]+\n/g, '\n')
        .replace(/\n{2,}/g, '\n')
        .trim()
      const frag = document.createElement('div')
      for (let i = 0; i < sel.rangeCount; i++) {
        frag.appendChild(sel.getRangeAt(i).cloneContents())
      }
      e.clipboardData.setData('text/plain', text)
      e.clipboardData.setData('text/html', frag.innerHTML)
      e.preventDefault()
    }
    el.addEventListener('copy', onCopy)
    return () => el.removeEventListener('copy', onCopy)
  }, [])

  // ---- 편집 진입: mammoth 로 편집용 HTML 준비(최초 1회) ----
  const enterEdit = useCallback(async () => {
    if (editHtml == null) {
      setPreparing(true)
      try {
        // mammoth 는 편집 진입 시에만 지연 로딩 (보기 속도에 영향 없음)
        const mammoth = (await import('mammoth/mammoth.browser.js')).default
        const { value } = await mammoth.convertToHtml({ arrayBuffer: buffer })
        setEditHtml(value && value.trim() ? value : '<p></p>')
      } catch (err) {
        console.error(err)
        setEditHtml('<p></p>')
        alert('편집용 변환 중 문제가 발생했습니다: ' + (err?.message || err))
      } finally {
        setPreparing(false)
      }
    }
    setEditMode(true)
  }, [buffer, editHtml])

  // 편집 HTML 이 준비되면 contentEditable 에 한 번 주입(이후 편집 내용 보존)
  useEffect(() => {
    if (editHtml != null && editRef.current && !editRef.current.dataset.filled) {
      editRef.current.innerHTML = editHtml
      editRef.current.dataset.filled = '1'
    }
  }, [editHtml])

  // contentEditable 선택 영역 저장/복원 (색상 피커가 포커스를 뺏어도 유지)
  const savedRange = useRef(null)
  const saveSel = useCallback(() => {
    const s = window.getSelection()
    if (s && s.rangeCount && editRef.current?.contains(s.anchorNode)) {
      savedRange.current = s.getRangeAt(0).cloneRange()
    }
  }, [])
  // 선택 영역을 인라인 스타일 span 으로 감싼다 (html-to-docx 가 run 단위 색상 인식)
  const wrapStyle = useCallback((prop, value) => {
    editRef.current?.focus()
    const sel = window.getSelection()
    const r = savedRange.current
    if (r) {
      sel.removeAllRanges()
      sel.addRange(r)
    }
    if (!sel.rangeCount || sel.isCollapsed) return
    const range = sel.getRangeAt(0)
    const span = document.createElement('span')
    span.style[prop] = value
    try {
      range.surroundContents(span)
    } catch {
      const frag = range.extractContents()
      span.appendChild(frag)
      range.insertNode(span)
    }
    sel.removeAllRanges()
    const nr = document.createRange()
    nr.selectNodeContents(span)
    sel.addRange(nr)
    savedRange.current = nr.cloneRange()
  }, [])

  const handleSave = useCallback(async () => {
    if (!editRef.current) return
    setSaving(true)
    try {
      const res = await exportDocxSave(
        editedName(name),
        editRef.current.innerHTML
      )
      if (res && res.appOnly) alert(res.error)
      else if (res && res.error) alert('저장 실패: ' + res.error)
      else if (res && res.saved) {
        setSaved(true)
        setTimeout(() => setSaved(false), 1800)
      }
    } catch (err) {
      alert('저장 실패: ' + (err?.message || err))
    } finally {
      setSaving(false)
    }
  }, [name])

  return (
    <div className="docx-outer">
      <div className="edit-toolbar">
        <button
          className={'tool-btn' + (editMode ? ' on' : '')}
          disabled={preparing}
          onClick={() => (editMode ? setEditMode(false) : enterEdit())}
        >
          {preparing ? '편집 준비 중…' : editMode ? '✏️ 편집 중' : '✏️ 편집'}
        </button>
        {editMode && (
          <>
            <button
              className="tool-btn primary"
              disabled={saving}
              onClick={handleSave}
            >
              {saving ? '저장 중…' : saved ? '✓ 저장됨' : '💾 다른 이름으로 저장'}
            </button>
            <span className="tool-sep" />
            <button
              className="tool-btn fmt"
              title="굵게"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('bold')
              }}
            >
              <b>B</b>
            </button>
            <button
              className="tool-btn fmt"
              title="기울임"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('italic')
              }}
            >
              <i>I</i>
            </button>
            <button
              className="tool-btn fmt"
              title="밑줄"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('underline')
              }}
            >
              <u>U</u>
            </button>
            <label
              className="tool-btn fmt color-btn"
              title="글자색"
              onMouseDown={saveSel}
            >
              <span style={{ color: '#c0392b' }}>가</span>
              <input
                type="color"
                onChange={(e) => wrapStyle('color', e.target.value)}
              />
            </label>
            <label
              className="tool-btn fmt color-btn"
              title="형광펜"
              onMouseDown={saveSel}
            >
              <span style={{ background: '#ffe08a', padding: '0 3px' }}>H</span>
              <input
                type="color"
                onChange={(e) => wrapStyle('backgroundColor', e.target.value)}
              />
            </label>
            <span className="tool-sep" />
            <button
              className="tool-btn fmt"
              title="왼쪽"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('justifyLeft')
              }}
            >
              ⬅
            </button>
            <button
              className="tool-btn fmt"
              title="가운데"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('justifyCenter')
              }}
            >
              ↔
            </button>
            <button
              className="tool-btn fmt"
              title="오른쪽"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('justifyRight')
              }}
            >
              ➡
            </button>
            <span className="tool-sep" />
            <select
              className="tool-select"
              title="글자 크기 (pt)"
              defaultValue=""
              onMouseDown={saveSel}
              onChange={(e) => {
                if (e.target.value) wrapStyle('fontSize', e.target.value + 'pt')
                e.target.value = ''
              }}
            >
              <option value="">크기</option>
              {[10, 11, 12, 14, 16, 18, 24, 32].map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </>
        )}
        {editMode ? (
          <span className="tool-hint">
            텍스트 수정 + 서식(B/I/U·색·형광·정렬·크기). 저장 시 정교한 서식은 일부 손실될 수 있어요.
          </span>
        ) : (
          <span className="tool-hint">보기 모드 — 서식이 충실히 표시됩니다.</span>
        )}
      </div>

      <div className="docx-scroll">
        {/* 보기 영역 */}
        {!editMode && status === 'loading' && (
          <div className="state-note">문서를 렌더링하는 중…</div>
        )}
        {!editMode && status === 'error' && (
          <div className="state-note error">문서를 열 수 없습니다: {error}</div>
        )}
        <div
          ref={containerRef}
          className="docx-host"
          style={{
            display: !editMode && status === 'done' ? 'block' : 'none',
            zoom,
          }}
        />

        {/* 편집 영역 (항상 마운트, 표시만 토글하여 편집 내용 유지) */}
        <div
          className="docx-edit-page"
          style={{ display: editMode ? 'block' : 'none', zoom }}
        >
          <div
            ref={editRef}
            className="docx-edit"
            contentEditable
            suppressContentEditableWarning
          />
        </div>
      </div>
    </div>
  )
}
