import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { saveBytes, editedName } from '../lib/save.js'

// 바이트 → 문자열 (UTF-8 우선, 깨지면 한글 EUC-KR 폴백)
function decodeText(buffer) {
  const bytes = new Uint8Array(buffer)
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
    return new TextDecoder('utf-8').decode(bytes.subarray(3))
  const u = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
  if (!u.includes('�')) return u
  try {
    return new TextDecoder('euc-kr').decode(bytes)
  } catch {
    return u
  }
}

// 마크다운 렌더 결과에서 위험 요소 제거(로컬 파일 대비 최소 방어)
function sanitize(html) {
  const div = document.createElement('div')
  div.innerHTML = html
  div
    .querySelectorAll('script,style,iframe,object,embed,link,meta,form')
    .forEach((e) => e.remove())
  div.querySelectorAll('*').forEach((el) => {
    ;[...el.attributes].forEach((a) => {
      if (
        /^on/i.test(a.name) ||
        (/^(href|src)$/i.test(a.name) && /^\s*javascript:/i.test(a.value))
      )
        el.removeAttribute(a.name)
    })
  })
  return div.innerHTML
}

export default function TextView({
  buffer,
  name = 'file.txt',
  kind = 'txt',
  zoom = 1,
}) {
  const original = useMemo(() => decodeText(buffer), [buffer])
  const [editMode, setEditMode] = useState(false)
  const [text, setText] = useState(original)
  const [dirty, setDirty] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [wrap, setWrap] = useState(true)
  const [html, setHtml] = useState('')
  const isMd = kind === 'md'

  useEffect(() => {
    setText(original)
    setDirty(false)
  }, [original])

  // 마크다운 렌더(보기 모드) — marked 지연 로딩
  useEffect(() => {
    let cancelled = false
    if (isMd && !editMode) {
      import('marked').then(({ marked }) => {
        if (!cancelled) setHtml(sanitize(marked.parse(text)))
      })
    }
    return () => {
      cancelled = true
    }
  }, [isMd, editMode, text])

  const handleSave = useCallback(async () => {
    setSaving(true)
    try {
      const bytes = new TextEncoder().encode(text)
      const res = await saveBytes(editedName(name), bytes)
      if (res.saved) {
        setDirty(false)
        setSaved(true)
        setTimeout(() => setSaved(false), 1800)
      } else if (res.error) alert('저장 실패: ' + res.error)
    } catch (err) {
      alert('저장 실패: ' + (err?.message || err))
    } finally {
      setSaving(false)
    }
  }, [text, name])

  return (
    <div className="text-wrap">
      <div className="edit-toolbar">
        <button
          className={'tool-btn' + (editMode ? ' on' : '')}
          onClick={() => setEditMode((v) => !v)}
          title="편집 모드 켜기/끄기"
        >
          {editMode ? '✏️ 편집 중' : '✏️ 편집'}
        </button>
        {editMode && (
          <button
            className="tool-btn primary"
            disabled={!dirty || saving}
            onClick={handleSave}
          >
            {saving ? '저장 중…' : saved ? '✓ 저장됨' : '💾 저장'}
          </button>
        )}
        {!isMd && !editMode && (
          <button
            className={'tool-btn' + (wrap ? ' on' : '')}
            onClick={() => setWrap((v) => !v)}
            title="줄바꿈(자동 넘김) 켜기/끄기"
          >
            ↩ 줄바꿈
          </button>
        )}
        <span className="tool-hint">
          {isMd
            ? editMode
              ? '마크다운 원문 편집'
              : '마크다운 미리보기'
            : `${kind.toUpperCase()} · 텍스트`}
        </span>
      </div>

      <div className="text-scroll">
        {editMode ? (
          <textarea
            className="text-edit"
            style={{ zoom }}
            value={text}
            onChange={(e) => {
              setText(e.target.value)
              setDirty(true)
            }}
            spellCheck={false}
          />
        ) : isMd ? (
          <div
            className="md-view"
            style={{ zoom }}
            dangerouslySetInnerHTML={{ __html: html }}
          />
        ) : (
          <pre
            className={'text-view' + (wrap ? ' wrap' : '')}
            style={{ zoom }}
          >
            {text}
          </pre>
        )}
      </div>
    </div>
  )
}
