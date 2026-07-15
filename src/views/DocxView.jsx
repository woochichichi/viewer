import React, { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'

// docx-preview: 스타일 / 표 / 이미지 보존.
// (mammoth 는 서식 손실이 커서 서식 보존 목적엔 docx-preview 사용)
export default function DocxView({ buffer }) {
  const containerRef = useRef(null)
  const [status, setStatus] = useState('loading') // loading | done | error
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false
    const el = containerRef.current
    if (!el) return
    el.innerHTML = ''
    setStatus('loading')
    setError('')

    // renderAsync 는 ArrayBuffer / Blob / Uint8Array 를 받는다.
    renderAsync(buffer, el, undefined, {
      className: 'docx',
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      breakPages: true,
      experimental: true,
      useBase64URL: true, // 이미지를 data URL 로 인라인 → 외부 요청 없음
    })
      .then(() => {
        if (!cancelled) setStatus('done')
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

  // 복사 시 문단(블록)마다 빈 줄이 하나씩 끼는 문제 보정.
  // 선택 영역이 문서 안일 때, 붙여넣기용 평문에서 연속 개행을 한 줄로 정리한다.
  // (서식 붙여넣기를 위해 HTML 은 원본 그대로 유지)
  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const onCopy = (e) => {
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return
      if (!el.contains(sel.anchorNode) && !el.contains(sel.focusNode)) return
      // 평문: 빈 줄(연속 개행)을 한 줄로, 각 줄 끝 공백 정리
      const text = sel
        .toString()
        .replace(/[ \t]+\n/g, '\n')
        .replace(/\n{2,}/g, '\n')
        .trim()
      // HTML: 선택 영역 원본 보존 (Word 등에 서식 유지 붙여넣기)
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

  return (
    <div className="docx-scroll">
      {status === 'loading' && (
        <div className="state-note">문서를 렌더링하는 중…</div>
      )}
      {status === 'error' && (
        <div className="state-note error">
          문서를 열 수 없습니다: {error}
        </div>
      )}
      <div
        ref={containerRef}
        className="docx-host"
        style={{ display: status === 'done' ? 'block' : 'none' }}
      />
    </div>
  )
}
