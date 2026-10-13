import React, { useCallback, useEffect, useRef, useState } from 'react'
import ExcelJS from 'exceljs/dist/exceljs.min.js'
import { saveBytes, editedName } from '../lib/save.js'
import { AlignIcon } from '../lib/icons.jsx'

// 1-indexed 열 인덱스 → 열 문자 (1→A, 27→AA)
function colLabel(c) {
  let s = ''
  while (c > 0) {
    const r = (c - 1) % 26
    s = String.fromCharCode(65 + r) + s
    c = Math.floor((c - 1) / 26)
  }
  return s
}
function decodeAddr(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a)
  if (!m) return { r: 1, c: 1 }
  let c = 0
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64)
  return { c, r: parseInt(m[2], 10) }
}

const THEME = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4',
  'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
]
function applyTint(hex, tint) {
  if (tint == null || tint === 0) return hex
  const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16))
  const f = (c) =>
    tint < 0 ? Math.round(c * (1 + tint)) : Math.round(c * (1 - tint) + 255 * tint)
  const h = (c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')
  return h(f(ch[0])) + h(f(ch[1])) + h(f(ch[2]))
}
function colorToCss(color) {
  if (!color) return null
  if (color.argb && color.argb.length >= 6) return '#' + color.argb.slice(-6)
  if (color.theme != null) {
    const base = THEME[color.theme]
    if (!base) return null
    return '#' + applyTint(base, color.tint)
  }
  return null
}
function widthToPx(w) {
  return typeof w === 'number' ? Math.round(w * 7 + 5) : null
}
function borderCss(b) {
  if (!b || !b.style) return null
  const w = b.style === 'medium' ? 2 : b.style === 'thick' ? 3 : 1
  return `${w}px solid ${colorToCss(b.color) || '#9aa4b2'}`
}

// 워크북 → 렌더 모델(시트 배열)
function sheetsFromWorkbook(wb) {
  return wb.worksheets.map((ws) => {
    const merges = ws.model.merges || []
    const covered = new Set()
    const spanMap = new Map()
    for (const m of merges) {
      const [s, e] = m.split(':')
      const sc = decodeAddr(s)
      const ec = decodeAddr(e)
      spanMap.set(`${sc.r},${sc.c}`, {
        rowSpan: ec.r - sc.r + 1,
        colSpan: ec.c - sc.c + 1,
      })
      for (let r = sc.r; r <= ec.r; r++)
        for (let c = sc.c; c <= ec.c; c++) {
          if (r === sc.r && c === sc.c) continue
          covered.add(`${r},${c}`)
        }
    }
    const colCount = Math.max(ws.columnCount, 1)
    const rowCount = Math.max(ws.rowCount, 1)
    const colWidths = []
    for (let c = 1; c <= colCount; c++) {
      const col = ws.getColumn(c)
      colWidths.push(col && col.hidden ? 0 : widthToPx(col && col.width))
    }
    const rows = []
    for (let r = 1; r <= rowCount; r++) {
      const row = ws.getRow(r)
      const cells = []
      for (let c = 1; c <= colCount; c++) {
        const key = `${r},${c}`
        if (covered.has(key)) {
          cells.push(null)
          continue
        }
        const cell = ws.getCell(r, c)
        const raw = cell.value
        let numeric = typeof raw === 'number'
        if (!numeric && raw && typeof raw === 'object' && typeof raw.result === 'number')
          numeric = true
        const text = cell.text != null ? String(cell.text) : ''
        const st = {}
        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
          const bg = colorToCss(cell.fill.fgColor)
          if (bg) st.backgroundColor = bg
        }
        const f = cell.font || {}
        if (f.bold) st.fontWeight = 'bold'
        if (f.italic) st.fontStyle = 'italic'
        const deco = []
        if (f.underline) deco.push('underline')
        if (f.strike) deco.push('line-through')
        if (deco.length) st.textDecoration = deco.join(' ')
        const fc = colorToCss(f.color)
        if (fc) st.color = fc
        if (f.size) st.fontSize = Math.round(f.size * 1.33) + 'px'
        if (f.name) st.fontFamily = `"${f.name}", inherit`
        const a = cell.alignment || {}
        if (a.horizontal && a.horizontal !== 'fill') st.textAlign = a.horizontal
        else if (numeric) st.textAlign = 'right'
        if (a.wrapText) st.whiteSpace = 'normal'
        const b = cell.border || {}
        if (borderCss(b.top)) st.borderTop = borderCss(b.top)
        if (borderCss(b.bottom)) st.borderBottom = borderCss(b.bottom)
        if (borderCss(b.left)) st.borderLeft = borderCss(b.left)
        if (borderCss(b.right)) st.borderRight = borderCss(b.right)
        const span = spanMap.get(key)
        cells.push({
          text,
          numeric,
          style: st,
          rowSpan: span?.rowSpan || 1,
          colSpan: span?.colSpan || 1,
        })
      }
      rows.push({
        cells,
        hidden: !!row.hidden,
        height: row.height ? Math.round(row.height * 1.33) : null,
      })
    }
    return { name: ws.name, colCount, rowCount, colWidths, rows }
  })
}

// CSV 바이트 → 문자열 (UTF-8 우선, 깨지면 한글 EUC-KR 로 재시도)
function decodeCsv(buffer) {
  const bytes = new Uint8Array(buffer)
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
    return new TextDecoder('utf-8').decode(bytes)
  const utf8 = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
  if (!utf8.includes('�')) return utf8
  try {
    return new TextDecoder('euc-kr').decode(bytes)
  } catch {
    return utf8
  }
}

async function buildModel(buffer, kind) {
  let xlsxBuf = buffer
  // .xls(구형 바이너리) / .csv → SheetJS 로 읽어 xlsx 로 변환(지연 로딩)
  if (kind === 'xls' || kind === 'csv') {
    const XLSX = await import('xlsx')
    const wbS =
      kind === 'csv'
        ? XLSX.read(decodeCsv(buffer), { type: 'string' })
        : XLSX.read(new Uint8Array(buffer), { type: 'array' })
    xlsxBuf = XLSX.write(wbS, { type: 'array', bookType: 'xlsx' })
  }
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(xlsxBuf)
  return { wb, sheets: sheetsFromWorkbook(wb) }
}

function applyEdit(ws, r, c, text) {
  const cell = ws.getCell(r, c)
  const t = text.trim()
  if (t === '') cell.value = null
  else if (/^-?\d+(\.\d+)?$/.test(t)) cell.value = Number(t)
  else cell.value = text
}

const toArgb = (hex) => 'FF' + hex.replace('#', '').slice(-6).toUpperCase()

// 대기 중 편집 키 (시트 이름에 공백이 있어도 안전하도록 객체 값에 좌표를 따로 보관)
const cellKey = (sheet, r, c) => `${sheet}\u0001${r}\u0001${c}`

// 대기 중인 서식 편집 → 화면 CSS (재렌더/시트 전환 후에도 편집 내용이 보이도록)
function pendingCss(st) {
  if (!st) return null
  const css = {}
  if (st.bold !== undefined) css.fontWeight = st.bold ? 'bold' : 'normal'
  if (st.color) css.color = st.color
  if (st.fill) css.backgroundColor = st.fill
  if (st.align) css.textAlign = st.align
  if (st.size) css.fontSize = Math.round(st.size * 1.33) + 'px'
  if (st.strike !== undefined) css.textDecoration = st.strike ? 'line-through' : 'none'
  return css
}

// 행/열 삽입·삭제 시 병합 범위 이동.
// ExcelJS 의 spliceRows/spliceColumns 는 병합 정보를 옮기지 않아(그대로 남거나 저장 시 사라짐)
// 병합을 모두 풀고 → 구조 변경 → 좌표를 보정해 다시 병합한다.
function spliceWithMerges(ws, op, at) {
  const ranges = (ws.model.merges || []).map((m) => {
    const [a, b] = m.split(':')
    const s = decodeAddr(a)
    const e = decodeAddr(b)
    return { r1: s.r, c1: s.c, r2: e.r, c2: e.c, ref: m }
  })
  for (const m of ranges) ws.unMergeCells(m.ref)
  const isRow = op === 'insRow' || op === 'delRow'
  const ins = op === 'insRow' || op === 'insCol'
  if (op === 'insRow') ws.spliceRows(at, 0, [])
  else if (op === 'delRow') ws.spliceRows(at, 1)
  else if (op === 'insCol') ws.spliceColumns(at, 0, [])
  else if (op === 'delCol') ws.spliceColumns(at, 1)
  const [k1, k2] = isRow ? ['r1', 'r2'] : ['c1', 'c2']
  for (const m of ranges) {
    if (ins) {
      if (m[k1] >= at) m[k1]++
      if (m[k2] >= at) m[k2]++
    } else {
      if (m[k2] < at) {
        // 영향 없음
      } else if (m[k1] > at) {
        m[k1]--
        m[k2]--
      } else m[k2]-- // 지운 줄이 병합 범위 안 → 범위 축소
    }
    if (m.r2 < m.r1 || m.c2 < m.c1) continue
    if (m.r1 === m.r2 && m.c1 === m.c2) continue // 한 칸만 남으면 병합 해제
    ws.mergeCells(m.r1, m.c1, m.r2, m.c2)
  }
}

const FONT_SIZES = [10, 11, 12, 14, 16, 18, 24, 32]

export default function XlsxView({
  buffer,
  name = 'sheet.xlsx',
  kind = 'xlsx',
  zoom = 1,
  onPrint,
}) {
  const [model, setModel] = useState(null)
  const [activeSheet, setActiveSheet] = useState(0)
  const [editMode, setEditMode] = useState(false)
  const [dirty, setDirty] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const editsRef = useRef(new Map()) // cellKey -> { sheet, r, c, text }
  const styleEditsRef = useRef(new Map()) // cellKey -> { sheet, r, c, st:{bold,color,fill,align,size,strike} }
  // 표 DOM 을 새로 만들기 위한 버전(시트 전환·행열 변경 시 이전 셀의 편집 흔적이 남지 않게)
  const [ver, setVer] = useState(0)
  // 표를 새로 그릴 때 보여줄 대기 편집 스냅숏. 입력 중 재렌더에서 React 가 셀 글자를
  // 덮어써 커서가 튀지 않도록, 표를 새로 만들 때(시트 전환)만 갱신한다.
  const snapRef = useRef({ edits: new Map(), styles: new Map() })
  const selRef = useRef(null) // { sheet, r, c, el }
  const rangeRef = useRef(null) // { sheet, r1, c1, r2, c2 }
  const gridRef = useRef(null)

  useEffect(() => {
    let cancelled = false
    setModel(null)
    editsRef.current = new Map()
    styleEditsRef.current = new Map()
    snapRef.current = { edits: new Map(), styles: new Map() }
    selRef.current = null
    rangeRef.current = null
    setDirty(false)
    buildModel(buffer, kind)
      .then((m) => !cancelled && setModel({ data: m }))
      .catch((e) => {
        console.error(e)
        if (!cancelled) setModel({ error: e?.message || String(e) })
      })
    return () => {
      cancelled = true
    }
  }, [buffer, kind])

  // 선택한 셀의 행·열 머리글 강조 (엑셀처럼 현재 위치 표시)
  const markHeads = (r, c) => {
    const g = gridRef.current
    if (!g) return
    g.querySelectorAll('th.hl').forEach((el) => el.classList.remove('hl'))
    g.querySelector(`th.col-head[data-c="${c}"]`)?.classList.add('hl')
    g.querySelector(`th.row-head[data-r="${r}"]`)?.classList.add('hl')
  }

  const tdAt = (r, c) =>
    gridRef.current?.querySelector(`td[data-r="${r}"][data-c="${c}"]`)

  const clearRangeHighlight = () => {
    gridRef.current
      ?.querySelectorAll('td.cell-sel')
      .forEach((el) => el.classList.remove('cell-sel'))
  }
  const highlightRange = (rr) => {
    clearRangeHighlight()
    if (!rr) return
    for (let r = Math.min(rr.r1, rr.r2); r <= Math.max(rr.r1, rr.r2); r++)
      for (let c = Math.min(rr.c1, rr.c2); c <= Math.max(rr.c1, rr.c2); c++)
        tdAt(r, c)?.classList.add('cell-sel')
  }

  // 셀 키보드: Enter 는 줄바꿈 대신 아래 칸으로(엑셀처럼). 셀 안 줄바꿈은 저장 시 사라져
  // 화면과 저장 결과가 달라지므로 막는다. Tab=오른쪽, Shift 와 함께 반대 방향.
  const onCellKeyDown = useCallback((e, r, c) => {
    // Ctrl+B/I/U: 브라우저 기본 서식(<b> 등)은 셀 서식으로 저장되지 않으므로 가로챈다
    if ((e.ctrlKey || e.metaKey) && !e.altKey) {
      const k = e.key.toLowerCase()
      if (k === 'b' || e.code === 'KeyB') {
        e.preventDefault()
        applyFmt({ bold: 'toggle' })
      } else if (k === 'i' || k === 'u' || e.code === 'KeyI' || e.code === 'KeyU') {
        e.preventDefault()
      }
      return
    }
    if (e.key !== 'Enter' && e.key !== 'Tab') return
    if (e.nativeEvent.isComposing) return // 한글 조합 중 Enter 는 조합 확정용
    e.preventDefault()
    const back = e.shiftKey
    const [dr, dc] = e.key === 'Enter' ? [back ? -1 : 1, 0] : [0, back ? -1 : 1]
    for (let i = 1; i < 500; i++) {
      const nr = r + dr * i
      const nc = c + dc * i
      if (nr < 1 || nc < 1) break
      const el = gridRef.current?.querySelector(`td[data-r="${nr}"][data-c="${nc}"]`)
      if (el) {
        el.focus()
        // 커서를 셀 끝으로
        const range = document.createRange()
        range.selectNodeContents(el)
        range.collapse(false)
        const sel = window.getSelection()
        sel.removeAllRanges()
        sel.addRange(range)
        el.scrollIntoView({ block: 'nearest', inline: 'nearest' })
        rangeRef.current = null
        clearRangeHighlight()
        return
      }
      if (!gridRef.current?.querySelector(`tr td[data-r="${nr}"]`) && dr) break
    }
  }, [])

  // 붙여넣기는 글자만 (엑셀/웹에서 복사한 표·서식 HTML 이 셀 안에 통째로 들어가는 것 방지)
  const onCellPaste = useCallback((e) => {
    e.preventDefault()
    const t = (e.clipboardData?.getData('text/plain') || '').replace(/\r?\n$/, '')
    document.execCommand('insertText', false, t.replace(/[\t\r\n]+/g, ' '))
  }, [])

  const onCellEdit = useCallback((sheetName, r, c, text) => {
    editsRef.current.set(cellKey(sheetName, r, c), { sheet: sheetName, r, c, text })
    setDirty(true)
  }, [])

  // 현재 선택(단일 셀 또는 범위)의 좌표 목록
  const targets = () => {
    const rr = rangeRef.current
    if (rr) {
      const list = []
      for (let r = Math.min(rr.r1, rr.r2); r <= Math.max(rr.r1, rr.r2); r++)
        for (let c = Math.min(rr.c1, rr.c2); c <= Math.max(rr.c1, rr.c2); c++)
          list.push({ sheet: rr.sheet, r, c })
      return list
    }
    const s = selRef.current
    return s ? [{ sheet: s.sheet, r: s.r, c: s.c }] : []
  }

  const applyFmt = useCallback((delta) => {
    const list = targets()
    if (!list.length) return
    // 토글류는 첫 셀 기준으로 on/off 결정
    let d = { ...delta }
    const first = tdAt(list[0].r, list[0].c)
    if (delta.bold === 'toggle') d = { bold: first?.style.fontWeight !== 'bold' }
    if (delta.strike === 'toggle')
      d = { strike: !(first?.style.textDecoration || '').includes('line-through') }
    for (const t of list) {
      const key = cellKey(t.sheet, t.r, t.c)
      const prev = styleEditsRef.current.get(key)
      styleEditsRef.current.set(key, { ...t, st: { ...(prev?.st || {}), ...d } })
      const el = tdAt(t.r, t.c)
      if (!el) continue
      const s = el.style
      if (d.bold !== undefined) s.fontWeight = d.bold ? 'bold' : 'normal'
      if (d.color) s.color = d.color
      if (d.fill) s.backgroundColor = d.fill
      if (d.align) s.textAlign = d.align
      if (d.size) s.fontSize = Math.round(d.size * 1.33) + 'px'
      if (d.strike !== undefined) s.textDecoration = d.strike ? 'line-through' : 'none'
    }
    setDirty(true)
  }, [])

  const flushToWb = (wb) => {
    for (const { sheet, r, c, text } of editsRef.current.values()) {
      const ws = wb.getWorksheet(sheet)
      if (ws) applyEdit(ws, r, c, text)
    }
    for (const { sheet, r, c, st } of styleEditsRef.current.values()) {
      const ws = wb.getWorksheet(sheet)
      if (!ws) continue
      const cell = ws.getCell(r, c)
      // 불러온 셀들은 같은 서식 객체를 공유한다 → cell.font = … 처럼 바로 바꾸면 같은 서식의
      // 다른 셀(시트 전체일 수도)까지 바뀐다. 이 셀 전용 style 사본을 만든 뒤 수정한다.
      const style = { ...(cell.style || {}) }
      if (st.bold !== undefined || st.color || st.size || st.strike !== undefined) {
        const font = { ...(style.font || {}) }
        if (st.bold !== undefined) font.bold = st.bold
        if (st.strike !== undefined) font.strike = st.strike
        if (st.color) font.color = { argb: toArgb(st.color) }
        if (st.size) font.size = st.size
        style.font = font
      }
      if (st.fill)
        style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(st.fill) } }
      if (st.align) style.alignment = { ...(style.alignment || {}), horizontal: st.align }
      cell.style = style
    }
  }

  const handleSave = useCallback(async () => {
    if (!model?.data) return
    setSaving(true)
    try {
      const wb = model.data.wb
      flushToWb(wb)
      // 값을 고쳤으면 수식 결과(캐시값)가 낡았으므로 엑셀이 열 때 다시 계산하도록
      wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true }
      const out = await wb.xlsx.writeBuffer()
      // xls/csv 로 연 파일도 편집 저장은 xlsx 로 (서식 보존)
      const outName = /\.(xls|csv)$/i.test(name)
        ? name.replace(/\.[^.]+$/, '') + '-편집.xlsx'
        : editedName(name)
      const res = await saveBytes(outName, new Uint8Array(out))
      if (res.saved) {
        setDirty(false)
        setSaved(true)
        setTimeout(() => setSaved(false), 1800)
        // 편집 목록은 지우지 않는다: 화면 모델은 원본 기준이라 시트 전환 시 편집값을 다시
        // 보여줘야 하고, 다음 저장 때 같은 값을 다시 반영해도 결과는 같다.
      } else if (res.error) alert('저장 실패: ' + res.error)
    } catch (err) {
      alert('저장 실패: ' + (err?.message || err))
    } finally {
      setSaving(false)
    }
  }, [model, name])

  // 행/열 삽입·삭제: 대기 편집을 워크북에 반영 후 구조 변경 → 모델 재생성
  // Ctrl+S (App 이 보냄): 이 탭이 활성이고 변경 있을 때 저장
  const rootRef = useRef(null)
  const saveRef = useRef(null)
  saveRef.current = () => {
    if (dirty && !saving) handleSave()
  }
  useEffect(() => {
    const h = () => rootRef.current?.closest('.view-pane.active') && saveRef.current()
    window.addEventListener('dv-save', h)
    return () => window.removeEventListener('dv-save', h)
  }, [])

  const structOp = useCallback(
    (op) => {
      const sel = selRef.current
      if (!model?.data || !sel) {
        alert('먼저 셀을 클릭해 위치를 정하세요.')
        return
      }
      const wb = model.data.wb
      flushToWb(wb)
      const ws = wb.getWorksheet(sel.sheet)
      if (!ws) return
      const at = op === 'insRow' ? sel.r + 1 : op === 'insCol' ? sel.c + 1 : op === 'delRow' ? sel.r : sel.c
      spliceWithMerges(ws, op, at)
      editsRef.current.clear()
      styleEditsRef.current.clear()
      snapRef.current = { edits: new Map(), styles: new Map() }
      selRef.current = null
      rangeRef.current = null
      setModel({ data: { wb, sheets: sheetsFromWorkbook(wb) } })
      setVer((v) => v + 1)
      setDirty(true)
    },
    [model]
  )

  if (!model)
    return (
      <div className="state-note loading" ref={rootRef}>
        스프레드시트를 여는 중…
      </div>
    )
  if (model.error)
    return (
      <div className="state-note error" ref={rootRef}>
        스프레드시트를 열 수 없습니다: {model.error}
        <br />
        <small>파일이 손상되었거나 암호가 걸린 문서일 수 있습니다.</small>
      </div>
    )

  const { sheets } = model.data
  if (!sheets.length)
    return (
      <div className="state-note error" ref={rootRef}>
        표시할 워크시트가 없습니다 (차트 시트만 있는 파일).
      </div>
    )
  const sheet = sheets[Math.min(activeSheet, sheets.length - 1)]

  return (
    <div className="xlsx-wrap" ref={rootRef}>
      <div className="edit-toolbar">
        <button
          className={'tool-btn' + (editMode ? ' on' : '')}
          onClick={() => setEditMode((v) => !v)}
          title={editMode ? '편집 끝내기 (편집 내용은 유지)' : '편집 모드로 바꾸기 (셀 클릭해 수정)'}
        >
          {editMode ? '✏️ 편집 중' : '✏️ 편집'}
        </button>
        <button
          className="tool-btn primary"
          disabled={!dirty || saving}
          onClick={handleSave}
          title={
            dirty
              ? `새 파일(이름-편집.xlsx)로 저장 (Ctrl+S)`
              : '변경 내용이 없습니다'
          }
        >
          {saving ? '저장 중…' : saved ? '✓ 저장됨' : '💾 저장'}
        </button>

        {editMode && (
          <>
            <span className="tool-sep" />
            <button
              className="tool-btn fmt"
              title="굵게 (Ctrl+B)"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ bold: 'toggle' })}
            >
              <b>B</b>
            </button>
            <button
              className="tool-btn fmt"
              title="취소선"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ strike: 'toggle' })}
            >
              <s>S</s>
            </button>
            <select
              className="tool-select"
              title="글자 크기"
              defaultValue=""
              onMouseDown={(e) => e.stopPropagation()}
              onChange={(e) => {
                if (e.target.value) applyFmt({ size: Number(e.target.value) })
                e.target.value = ''
              }}
            >
              <option value="">크기</option>
              {FONT_SIZES.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
            <label className="tool-btn fmt color-btn" title="글자색 (선택한 셀)">
              <span style={{ color: '#c0392b' }}>가</span>
              <input type="color" onChange={(e) => applyFmt({ color: e.target.value })} />
            </label>
            <label className="tool-btn fmt color-btn" title="채우기 색 (선택한 셀)">
              <span style={{ background: '#ffe08a', padding: '0 3px' }}>■</span>
              <input type="color" onChange={(e) => applyFmt({ fill: e.target.value })} />
            </label>
            <span className="tool-sep" />
            <button
              className="tool-btn fmt"
              title="왼쪽 정렬"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ align: 'left' })}
            >
              <AlignIcon dir="left" />
            </button>
            <button
              className="tool-btn fmt"
              title="가운데 정렬"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ align: 'center' })}
            >
              <AlignIcon dir="center" />
            </button>
            <button
              className="tool-btn fmt"
              title="오른쪽 정렬"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ align: 'right' })}
            >
              <AlignIcon dir="right" />
            </button>
            <span className="tool-sep" />
            <button className="tool-btn fmt" title="선택한 셀 아래에 행 삽입" onClick={() => structOp('insRow')}>
              +행
            </button>
            <button className="tool-btn fmt" title="선택한 셀의 행 삭제" onClick={() => structOp('delRow')}>
              −행
            </button>
            <button className="tool-btn fmt" title="선택한 셀 오른쪽에 열 삽입" onClick={() => structOp('insCol')}>
              +열
            </button>
            <button className="tool-btn fmt" title="선택한 셀의 열 삭제" onClick={() => structOp('delCol')}>
              −열
            </button>
          </>
        )}
        {editMode ? (
          <span
            className="tool-hint"
            title="셀 클릭=값 수정 · Enter/Tab=다음 셀 · Shift+클릭=범위 선택 · 서식/행열 버튼 적용"
          >
            셀 클릭=값 수정 · Enter/Tab=다음 셀 · Shift+클릭=범위 선택
          </span>
        ) : (
          <span className="tool-hint">
            {dirty ? '저장하지 않은 변경이 있습니다.' : `${sheet.rowCount}행 × ${sheet.colCount}열`}
          </span>
        )}
        <button className="tool-btn print-btn" onClick={onPrint} title="인쇄 (Ctrl+P)">
          🖨 인쇄
        </button>
      </div>

      <div className="xlsx-grid-scroll" ref={gridRef}>
        <table key={`${ver}:${sheet.name}`} className="xlsx-table" style={{ zoom }}>
          <colgroup>
            <col className="head-col" style={{ width: 44 }} />
            {sheet.colWidths.map((w, i) => (
              <col key={i} style={w != null ? { width: w } : undefined} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="corner" />
              {Array.from({ length: sheet.colCount }, (_, i) => (
                <th key={i} className="col-head" data-c={i + 1}>
                  {colLabel(i + 1)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.rows.map((row, ri) => (
              <tr
                key={ri}
                style={
                  row.hidden
                    ? { display: 'none' }
                    : row.height
                      ? { height: row.height }
                      : undefined
                }
              >
                <th className="row-head" data-r={ri + 1}>
                  {ri + 1}
                </th>
                {row.cells.map((cell, ci) => {
                  if (cell === null) return null
                  const k = cellKey(sheet.name, ri + 1, ci + 1)
                  const pend = snapRef.current.edits.get(k)
                  const pst = pendingCss(snapRef.current.styles.get(k)?.st)
                  return (
                    <td
                      key={ci}
                      data-r={ri + 1}
                      data-c={ci + 1}
                      className={editMode ? 'editable' : undefined}
                      style={pst ? { ...cell.style, ...pst } : cell.style}
                      rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                      colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                      contentEditable={editMode}
                      suppressContentEditableWarning
                      onMouseDown={
                        editMode
                          ? (e) => {
                              if (e.shiftKey && selRef.current) {
                                e.preventDefault()
                                rangeRef.current = {
                                  sheet: sheet.name,
                                  r1: selRef.current.r,
                                  c1: selRef.current.c,
                                  r2: ri + 1,
                                  c2: ci + 1,
                                }
                                highlightRange(rangeRef.current)
                              } else {
                                rangeRef.current = null
                                clearRangeHighlight()
                              }
                            }
                          : undefined
                      }
                      onFocus={
                        editMode
                          ? (e) => {
                              markHeads(ri + 1, ci + 1)
                              selRef.current = {
                                sheet: sheet.name,
                                r: ri + 1,
                                c: ci + 1,
                                el: e.currentTarget,
                              }
                            }
                          : undefined
                      }
                      onKeyDown={editMode ? (e) => onCellKeyDown(e, ri + 1, ci + 1) : undefined}
                      onPaste={editMode ? onCellPaste : undefined}
                      onInput={
                        editMode
                          ? (e) =>
                              onCellEdit(sheet.name, ri + 1, ci + 1, e.currentTarget.textContent)
                          : undefined
                      }
                    >
                      {pend ? pend.text : cell.text}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="sheet-tabs">
        {sheets.map((s, i) => (
          <button
            key={i}
            className={'sheet-tab' + (i === activeSheet ? ' active' : '')}
            onClick={() => {
              if (i === activeSheet) return
              snapRef.current = {
                edits: new Map(editsRef.current),
                styles: new Map(styleEditsRef.current),
              }
              setActiveSheet(i)
              selRef.current = null
              rangeRef.current = null
            }}
            title={s.name}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
