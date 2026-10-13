import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { saveBytes, editedName } from '../lib/save.js'

// 바이트 → { text, enc, bom, crlf }
// UTF-8(BOM) / UTF-16(BOM, 윈도우 메모장 "유니코드") 우선, 깨지면 한글 EUC-KR(CP949) 폴백.
// enc·bom·crlf 는 저장 시 원래 형식 그대로 되돌리기 위해 기억한다.
function decodeText(buffer) {
  const bytes = new Uint8Array(buffer)
  let text
  let enc = 'utf-8'
  let bom = false
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    text = new TextDecoder('utf-8').decode(bytes.subarray(3))
    bom = true
  } else if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    text = new TextDecoder('utf-16le').decode(bytes.subarray(2))
    enc = 'utf-16le'
    bom = true
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    text = new TextDecoder('utf-16be').decode(bytes.subarray(2))
    enc = 'utf-16be'
    bom = true
  } else {
    text = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
    if (text.includes('\uFFFD')) {
      try {
        text = new TextDecoder('euc-kr').decode(bytes)
        enc = 'euc-kr'
      } catch {}
    }
  }
  return { text, enc, bom, crlf: text.includes('\r\n') }
}

// CP949(EUC-KR) 역변환표: TextDecoder 로 모든 2바이트 조합을 풀어 만든다(외부 표 불필요).
let cp949Map = null
function getCp949Map() {
  if (cp949Map) return cp949Map
  cp949Map = new Map()
  const dec = new TextDecoder('euc-kr')
  const pair = new Uint8Array(2)
  for (let a = 0x81; a <= 0xfe; a++)
    for (let b = 0x41; b <= 0xfe; b++) {
      pair[0] = a
      pair[1] = b
      const ch = dec.decode(pair)
      if (ch.length === 1 && ch !== '\uFFFD' && !cp949Map.has(ch))
        cp949Map.set(ch, (a << 8) | b)
    }
  return cp949Map
}

// 문자열 → 원래 인코딩 바이트. CP949 로 표현 못 하는 글자(이모지 등)가 있으면 null.
function encodeText(text, { enc, bom }) {
  if (enc === 'utf-16le' || enc === 'utf-16be') {
    const out = new Uint8Array(2 + text.length * 2)
    const le = enc === 'utf-16le'
    out[0] = le ? 0xff : 0xfe
    out[1] = le ? 0xfe : 0xff
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i)
      out[2 + i * 2] = le ? c & 0xff : c >> 8
      out[3 + i * 2] = le ? c >> 8 : c & 0xff
    }
    return out
  }
  if (enc === 'euc-kr') {
    const map = getCp949Map()
    const out = []
    for (const ch of text) {
      const c = ch.charCodeAt(0)
      if (c < 0x80) out.push(c)
      else {
        const v = map.get(ch)
        if (v == null) return null
        out.push(v >> 8, v & 0xff)
      }
    }
    return new Uint8Array(out)
  }
  const u8 = new TextEncoder().encode(text)
  if (!bom) return u8
  const out = new Uint8Array(u8.length + 3)
  out.set([0xef, 0xbb, 0xbf])
  out.set(u8, 3)
  return out
}

// 마크다운 렌더 결과에서 위험 요소 제거(로컬 파일 대비 방어)
// DOMParser 로 파싱 → 스크립트 실행·이미지 로딩이 일어나지 않는 비활성 문서에서 정리.
const BAD_TAGS =
  'script,style,iframe,frame,frameset,object,embed,applet,link,meta,base,form,' +
  'svg,math,template,noscript'
const URL_ATTRS = /^(href|src|xlink:href|action|formaction|data|poster|background|cite|srcset|ping)$/i
function safeUrl(value, attr, tag) {
  // 브라우저는 URL 의 공백·제어문자를 무시하므로 지운 뒤 스킴 검사 (java\tscript: 등 우회 차단)
  const v = String(value).replace(/[\u0000-\u0020\u007f-\u009f]/g, '').toLowerCase()
  if (/^(javascript|vbscript|livescript):/.test(v)) return false
  if (v.startsWith('data:')) return tag === 'IMG' && /^src$/i.test(attr) && v.startsWith('data:image/')
  return true
}
function sanitize(html) {
  const doc = new DOMParser().parseFromString(
    '<!doctype html><body>' + html,
    'text/html'
  )
  const body = doc.body
  body.querySelectorAll(BAD_TAGS).forEach((e) => e.remove())
  body.querySelectorAll('*').forEach((el) => {
    ;[...el.attributes].forEach((a) => {
      if (
        /^on/i.test(a.name) ||
        /^(srcdoc|style)$/i.test(a.name) ||
        (URL_ATTRS.test(a.name) && !safeUrl(a.value, a.name, el.tagName))
      )
        el.removeAttribute(a.name)
    })
  })
  return body.innerHTML
}

export default function TextView({
  buffer,
  name = 'file.txt',
  kind = 'txt',
  zoom = 1,
  onPrint,
}) {
  const decoded = useMemo(() => decodeText(buffer), [buffer])
  const original = decoded.text
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
      // 원래 줄바꿈(CRLF)·인코딩 유지 — textarea 는 줄바꿈을 \n 으로 바꾸므로 되돌린다.
      // (.bat 은 LF 만 있으면 cmd 에서 레이블/goto 가 깨지고, CP949 → UTF-8 이면 한글이 깨짐)
      const body = decoded.crlf ? text.replace(/\r?\n/g, '\r\n') : text
      let bytes = encodeText(body, decoded)
      if (!bytes) {
        if (
          !confirm(
            '원래 인코딩(EUC-KR)으로 저장할 수 없는 글자(이모지 등)가 있습니다.\n' +
              'UTF-8 로 저장할까요?'
          )
        )
          return
        bytes = new TextEncoder().encode(body)
      }
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
  }, [text, name, decoded])

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
        <button className="tool-btn print-btn" onClick={onPrint} title="인쇄 (Ctrl+P)">
          🖨 인쇄
        </button>
      </div>

      <div className="text-scroll">
        {/* textarea 는 보이는 부분만 인쇄되므로 인쇄용 사본을 따로 둔다 */}
        {editMode && <pre className="text-view print-only">{text}</pre>}
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
