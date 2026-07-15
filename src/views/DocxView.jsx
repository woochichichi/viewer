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
