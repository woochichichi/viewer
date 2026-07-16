import React, { useCallback, useEffect, useRef, useState } from 'react'
import ExcelJS from 'exceljs/dist/exceljs.min.js'
import { saveBytes, editedName } from '../lib/save.js'

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

const FONT_SIZES = [10, 11, 12, 14, 16, 18, 24, 32]

export default function XlsxView({
  buffer,
  name = 'sheet.xlsx',
  kind = 'xlsx',
  zoom = 1,
}) {
  const [model, setModel] = useState(null)
  const [activeSheet, setActiveSheet] = useState(0)
  const [editMode, setEditMode] = useState(false)
  const [dirty, setDirty] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const editsRef = useRef(new Map()) // "sheet r c" -> 문자열
  const styleEditsRef = useRef(new Map()) // "sheet r c" -> {bold,color,fill,align,size,strike}
  const selRef = useRef(null) // { sheet, r, c, el }
  const rangeRef = useRef(null) // { sheet, r1, c1, r2, c2 }
  const gridRef = useRef(null)

  useEffect(() => {
    let cancelled = false
    setModel(null)
    editsRef.current = new Map()
    styleEditsRef.current = new Map()
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

  const onCellEdit = useCallback((sheetName, r, c, text) => {
    editsRef.current.set(`${sheetName} ${r} ${c}`, text)
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
      const key = `${t.sheet} ${t.r} ${t.c}`
      styleEditsRef.current.set(key, { ...(styleEditsRef.current.get(key) || {}), ...d })
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
    for (const [key, text] of editsRef.current) {
      const [sheetName, rs, cs] = key.split(' ')
      const ws = wb.getWorksheet(sheetName)
      if (ws) applyEdit(ws, Number(rs), Number(cs), text)
    }
    for (const [key, st] of styleEditsRef.current) {
      const [sheetName, rs, cs] = key.split(' ')
      const ws = wb.getWorksheet(sheetName)
      if (!ws) continue
      const cell = ws.getCell(Number(rs), Number(cs))
      if (st.bold !== undefined || st.color || st.size || st.strike !== undefined) {
        const font = { ...(cell.font || {}) }
        if (st.bold !== undefined) font.bold = st.bold
        if (st.strike !== undefined) font.strike = st.strike
        if (st.color) font.color = { argb: toArgb(st.color) }
        if (st.size) font.size = st.size
        cell.font = font
      }
      if (st.fill)
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(st.fill) } }
      if (st.align) cell.alignment = { ...(cell.alignment || {}), horizontal: st.align }
    }
  }

  const handleSave = useCallback(async () => {
    if (!model?.data) return
    setSaving(true)
    try {
      const wb = model.data.wb
      flushToWb(wb)
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
        editsRef.current.clear()
        styleEditsRef.current.clear()
      } else if (res.error) alert('저장 실패: ' + res.error)
    } catch (err) {
      alert('저장 실패: ' + (err?.message || err))
    } finally {
      setSaving(false)
    }
  }, [model, name])

  // 행/열 삽입·삭제: 대기 편집을 워크북에 반영 후 구조 변경 → 모델 재생성
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
      if (op === 'insRow') ws.spliceRows(sel.r + 1, 0, [])
      else if (op === 'delRow') ws.spliceRows(sel.r, 1)
      else if (op === 'insCol') ws.spliceColumns(sel.c + 1, 0, [])
      else if (op === 'delCol') ws.spliceColumns(sel.c, 1)
      editsRef.current.clear()
      styleEditsRef.current.clear()
      selRef.current = null
      rangeRef.current = null
      setModel({ data: { wb, sheets: sheetsFromWorkbook(wb) } })
      setDirty(true)
    },
    [model]
  )

  if (!model) return <div className="state-note">스프레드시트를 여는 중…</div>
  if (model.error)
    return (
      <div className="state-note error">스프레드시트를 열 수 없습니다: {model.error}</div>
    )

  const { sheets } = model.data
  const sheet = sheets[Math.min(activeSheet, sheets.length - 1)]

  return (
    <div className="xlsx-wrap">
      <div className="edit-toolbar">
        <button
          className={'tool-btn' + (editMode ? ' on' : '')}
          onClick={() => setEditMode((v) => !v)}
          title="편집 모드 켜기/끄기"
        >
          {editMode ? '✏️ 편집 중' : '✏️ 편집'}
        </button>
        <button
          className="tool-btn primary"
          disabled={!dirty || saving}
          onClick={handleSave}
          title="다른 이름으로 저장 (Ctrl+S 아님, 버튼 클릭)"
        >
          {saving ? '저장 중…' : saved ? '✓ 저장됨' : '💾 저장'}
        </button>

        {editMode && (
          <>
            <span className="tool-sep" />
            <button
              className="tool-btn fmt"
              title="굵게"
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
            <label className="tool-btn fmt color-btn" title="글자색">
              <span style={{ color: '#c0392b' }}>가</span>
              <input type="color" onChange={(e) => applyFmt({ color: e.target.value })} />
            </label>
            <label className="tool-btn fmt color-btn" title="채우기색">
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
              ⬅
            </button>
            <button
              className="tool-btn fmt"
              title="가운데 정렬"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ align: 'center' })}
            >
              ↔
            </button>
            <button
              className="tool-btn fmt"
              title="오른쪽 정렬"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyFmt({ align: 'right' })}
            >
              ➡
            </button>
            <span className="tool-sep" />
            <button className="tool-btn fmt" title="아래에 행 추가" onClick={() => structOp('insRow')}>
              ⊞행
            </button>
            <button className="tool-btn fmt" title="현재 행 삭제" onClick={() => structOp('delRow')}>
              ⊟행
            </button>
            <button className="tool-btn fmt" title="오른쪽에 열 추가" onClick={() => structOp('insCol')}>
              ⊞열
            </button>
            <button className="tool-btn fmt" title="현재 열 삭제" onClick={() => structOp('delCol')}>
              ⊟열
            </button>
          </>
        )}
        {editMode ? (
          <span className="tool-hint">
            셀 클릭=값 수정 · Shift+클릭=범위 선택 · 서식/행열 버튼 적용
          </span>
        ) : (
          dirty && <span className="tool-hint">저장하지 않은 변경이 있습니다.</span>
        )}
      </div>

      <div className="xlsx-grid-scroll" ref={gridRef}>
        <table className="xlsx-table" style={{ zoom }}>
          <colgroup>
            <col style={{ width: 44 }} />
            {sheet.colWidths.map((w, i) => (
              <col key={i} style={w != null ? { width: w } : undefined} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="corner" />
              {Array.from({ length: sheet.colCount }, (_, i) => (
                <th key={i} className="col-head">
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
                <th className="row-head">{ri + 1}</th>
                {row.cells.map((cell, ci) =>
                  cell === null ? null : (
                    <td
                      key={ci}
                      data-r={ri + 1}
                      data-c={ci + 1}
                      className={editMode ? 'editable' : undefined}
                      style={cell.style}
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
                              selRef.current = {
                                sheet: sheet.name,
                                r: ri + 1,
                                c: ci + 1,
                                el: e.currentTarget,
                              }
                            }
                          : undefined
                      }
                      onInput={
                        editMode
                          ? (e) =>
                              onCellEdit(sheet.name, ri + 1, ci + 1, e.currentTarget.textContent)
                          : undefined
                      }
                    >
                      {cell.text}
                    </td>
                  )
                )}
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
