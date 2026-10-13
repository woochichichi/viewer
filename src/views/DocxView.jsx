import React, { useCallback, useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { exportDocxSave, editedName } from '../lib/save.js'
import { AlignIcon } from '../lib/icons.jsx'
import { annotateDocx, paginate } from '../lib/paginate.js'

// 보기: docx-preview(서식 충실). 편집: mammoth 로 단순 HTML 화 → 편집 → docx 저장.
export default function DocxView({
  buffer,
  name = 'document.docx',
  docId = 0,
  zoom = 1,
  onAutoFit,
  onPrint,
}) {
  const containerRef = useRef(null) // docx-preview 보기 영역
  const scrollRef = useRef(null) // 스크롤 컨테이너
  const editRef = useRef(null) // 편집(contentEditable) 영역
  const [status, setStatus] = useState('loading') // loading | layout(페이지 나누는 중) | done | error
  const [error, setError] = useState('')
  const [editMode, setEditMode] = useState(false)
  const [editHtml, setEditHtml] = useState(null) // mammoth 결과(최초 1회)
  const [preparing, setPreparing] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)

  // A4 페이지를 창 너비에 맞춘다 ('페이지 맞춤' 버튼. 열 때는 App 이 기본 90%)
  const fitToWidth = useCallback(() => {
    const host = containerRef.current
    const scroll = scrollRef.current
    if (!host || !scroll) return
    const page = host.querySelector('.docx-wrapper > section')
    if (!page || !page.offsetWidth) return
    const avail = scroll.clientWidth - 48 // 좌우 패딩
    onAutoFit?.(avail / page.offsetWidth)
  }, [onAutoFit])

  // ---- 보기: docx-preview 렌더 ----
  useEffect(() => {
    let cancelled = false
    const el = containerRef.current
    if (!el) return
    el.innerHTML = ''
    setStatus('loading')
    setError('')
    renderAsync(buffer, el, undefined, {
      // 문서마다 고유 접두사: 여러 docx 를 탭으로 열어 두면 docx-preview 가 넣는 <style> 이
      // 전역으로 겹쳐 서로의 스타일(Normal·제목 등)을 덮어쓰므로 클래스 이름을 분리한다.
      className: 'docx' + docId,
      inWrapper: true,
      breakPages: true,
      // 워드가 마지막 저장 때 기록한 페이지 위치(lastRenderedPageBreak)는 쓰지 않는다.
      // 자동 페이지 나눔(paginate)이 실제 렌더 높이로 나누므로, 이 표시까지 쓰면 글꼴 차이로
      // 이중으로 나뉘어 몇 줄짜리 페이지가 생긴다. 명시적 페이지/구역 나누기는 그대로 유지.
      ignoreLastRenderedPageBreak: true,
      experimental: true,
      useBase64URL: true,
    })
      .then(async (doc) => {
        if (cancelled) return
        // 공통 CSS(그림자·인쇄 페이지 나눔)가 적용되도록 래퍼에 공통 클래스 부여
        const wrapper = el.querySelector(`.docx${docId}-wrapper`)
        wrapper?.classList.add('docx-wrapper')
        await annotateDocx(wrapper, doc, 'docx' + docId)
        if (cancelled) return
        // 화면에 보이게(visibility:hidden) 배치한 뒤 페이지를 나누고 표시한다(아래 effect)
        paginatedRef.current = false
        setStatus('layout')
      })
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

  // ---- 자동 페이지 나눔 ----
  // 측정은 100% 배율·화면에 보이는 상태에서만 가능 → 숨겨진 탭/편집 모드면 보일 때까지 미룬다.
  // 글꼴·그림이 늦게 로드되면 다시 나눈다(dirty → 보일 때 재실행).
  const paginatedRef = useRef(false)
  const dirtyRef = useRef(false)
  const runPaginate = useCallback(() => {
    const el = containerRef.current
    const wrapper = el?.querySelector('.docx-wrapper')
    if (!wrapper || !el.getClientRects().length) return false
    const prevZoom = el.style.zoom
    el.style.zoom = '1' // 배율과 무관하게 실제(100%) 크기로 측정
    try {
      paginate(wrapper)
    } catch (err) {
      console.error('paginate', err)
    } finally {
      el.style.zoom = prevZoom
    }
    dirtyRef.current = false
    return true
  }, [])

  // 최초: 글꼴·그림 로드를 기다린 뒤 나누고 표시
  useEffect(() => {
    if (status !== 'layout' || editMode) return
    const el = containerRef.current
    if (!el) return
    let stop = false
    let busy = false
    let ro = null
    const go = async () => {
      if (stop || busy || !el.getClientRects().length) return // 숨김 → ResizeObserver 가 다시 부름
      busy = true
      ro?.disconnect()
      ro = null
      const imgs = [...el.querySelectorAll('img')].filter((i) => !i.complete)
      const wait = Promise.all([
        document.fonts?.ready,
        ...imgs.map((i) => i.decode?.().catch(() => {})),
      ])
      await Promise.race([wait, new Promise((r) => setTimeout(r, 3000))])
      busy = false
      if (stop) return
      if (!runPaginate()) {
        ro = new ResizeObserver(go)
        ro.observe(el)
        return
      }
      paginatedRef.current = true
      setStatus('done')
    }
    ro = new ResizeObserver(go)
    ro.observe(el)
    go()
    return () => {
      stop = true
      ro?.disconnect()
    }
  }, [status, editMode, runPaginate])

  // 이후: 글꼴/그림이 늦게 로드되면 다시 나눔 (숨김 상태면 보일 때)
  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    let t = null
    const schedule = () => {
      if (!paginatedRef.current) return
      dirtyRef.current = true
      clearTimeout(t)
      t = setTimeout(() => dirtyRef.current && runPaginate(), 150)
    }
    // 나눔 때 복제한 머리말/꼬리말·반복 머리글 행의 그림은 제외(다시 나누면 또 복제 → 무한 반복)
    const CLONED = '[data-dv-page="cont"] > header, [data-dv-page="cont"] > footer, [data-dv-rephead]'
    const onLoad = (e) =>
      e.target?.tagName === 'IMG' && !e.target.closest(CLONED) && schedule()
    el.addEventListener('load', onLoad, true) // load 는 버블링 안 됨 → 캡처
    document.fonts?.addEventListener?.('loadingdone', schedule)
    const ro = new ResizeObserver(() => {
      if (dirtyRef.current && el.getClientRects().length) runPaginate()
    })
    ro.observe(el)
    return () => {
      clearTimeout(t)
      el.removeEventListener('load', onLoad, true)
      document.fonts?.removeEventListener?.('loadingdone', schedule)
      ro.disconnect()
    }
  }, [runPaginate])

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
    if (!editRef.current || saving) return
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
  }, [name, saving])

  // Ctrl+S (App 이 보냄): 이 탭이 활성이고 편집 중일 때만 저장
  const rootRef = useRef(null)
  const saveRef = useRef(null)
  saveRef.current = () => {
    if (editMode) handleSave() // 보기 모드는 바뀐 내용이 없으므로 무시
  }
  useEffect(() => {
    const h = () => rootRef.current?.closest('.view-pane.active') && saveRef.current()
    window.addEventListener('dv-save', h)
    return () => window.removeEventListener('dv-save', h)
  }, [])

  return (
    <div className="docx-outer" ref={rootRef}>
      <div className="edit-toolbar">
        <button
          className={'tool-btn' + (editMode ? ' on' : '')}
          disabled={preparing}
          onClick={() => (editMode ? setEditMode(false) : enterEdit())}
          title={editMode ? '편집 끝내고 원래 보기로 (편집 내용은 유지)' : '편집 모드로 바꾸기'}
        >
          {preparing ? '편집 준비 중…' : editMode ? '✏️ 편집 중' : '✏️ 편집'}
        </button>
        {editMode && (
          <>
            <button
              className="tool-btn primary"
              disabled={saving}
              onClick={handleSave}
              title="새 파일(이름-편집.docx)로 저장 (Ctrl+S)"
            >
              {saving ? '저장 중…' : saved ? '✓ 저장됨' : '💾 다른 이름으로 저장'}
            </button>
            <span className="tool-sep" />
            <button
              className="tool-btn fmt"
              title="굵게 (Ctrl+B)"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('bold')
              }}
            >
              <b>B</b>
            </button>
            <button
              className="tool-btn fmt"
              title="기울임 (Ctrl+I)"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('italic')
              }}
            >
              <i>I</i>
            </button>
            <button
              className="tool-btn fmt"
              title="밑줄 (Ctrl+U)"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('underline')
              }}
            >
              <u>U</u>
            </button>
            <label
              className="tool-btn fmt color-btn"
              title="글자색 (글자를 선택한 뒤)"
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
              title="형광펜 (글자를 선택한 뒤)"
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
              title="왼쪽 맞춤"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('justifyLeft')
              }}
            >
              <AlignIcon dir="left" />
            </button>
            <button
              className="tool-btn fmt"
              title="가운데 맞춤"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('justifyCenter')
              }}
            >
              <AlignIcon dir="center" />
            </button>
            <button
              className="tool-btn fmt"
              title="오른쪽 맞춤"
              onMouseDown={(e) => {
                e.preventDefault()
                document.execCommand('justifyRight')
              }}
            >
              <AlignIcon dir="right" />
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
          <span
            className="tool-hint"
            title="텍스트 수정 + 서식(B/I/U·색·형광·정렬·크기). 저장 시 정교한 서식은 일부 손실될 수 있어요."
          >
            텍스트 수정 + 서식(B/I/U·색·형광·정렬·크기). 저장 시 정교한 서식은 일부 손실될 수 있어요.
          </span>
        ) : (
          <>
            <button
              className="tool-btn"
              onClick={fitToWidth}
              title="A4 한 페이지를 창 너비에 맞춤"
            >
              ⤢ 페이지 맞춤
            </button>
            <span className="tool-hint" title="보기 모드 — 서식이 충실히 표시됩니다. 고치려면 [편집]">
              보기 모드 — 서식이 충실히 표시됩니다.
            </span>
          </>
        )}
        <button className="tool-btn print-btn" onClick={onPrint} title="인쇄 (Ctrl+P)">
          🖨 인쇄
        </button>
      </div>

      <div className="docx-scroll" ref={scrollRef}>
        {/* 보기 영역 */}
        {!editMode && (status === 'loading' || status === 'layout') && (
          <div className="state-note loading">문서를 렌더링하는 중…</div>
        )}
        {!editMode && status === 'error' && (
          <div className="state-note error">문서를 열 수 없습니다: {error}</div>
        )}
        <div
          ref={containerRef}
          className="docx-host"
          style={{
            display: !editMode && (status === 'done' || status === 'layout') ? 'block' : 'none',
            // 페이지 나누는 동안은 배치만 하고 감춘다(긴 한 장이 번쩍이지 않게)
            visibility: status === 'layout' ? 'hidden' : undefined,
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
