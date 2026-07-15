import React, { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

// 열 인덱스 → 스프레드시트 열 문자 (0→A, 26→AA)
function colLabel(n) {
  let s = ''
  n += 1
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

// !cols 항목에서 픽셀 너비 추정
function colPx(colInfo) {
  if (!colInfo) return null
  if (colInfo.hidden) return 0
  if (typeof colInfo.wpx === 'number') return Math.round(colInfo.wpx)
  if (typeof colInfo.wch === 'number') return Math.round(colInfo.wch * 7 + 5)
  return null
}

// 워크북을 화면 렌더링용 데이터로 변환
function buildModel(buffer) {
  // 수식은 결과값(.w / .v)으로 표시. cellFormula:false 로 수식 문자열 파싱 자체를 끔.
  const wb = XLSX.read(buffer, {
    type: 'array',
    cellFormula: false,
    cellStyles: true,
    cellDates: true,
    cellNF: true,
  })

  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name]
    const ref = ws['!ref'] || 'A1:A1'
    const range = XLSX.utils.decode_range(ref)
    const merges = ws['!merges'] || []

    // 병합에서 "가려지는" 좌상단 이외의 셀 집합 + 좌상단 셀의 span 정보
    const covered = new Set()
    const spanMap = new Map() // "r,c" -> { rowSpan, colSpan }
    for (const m of merges) {
      const rs = m.s.r
      const cs = m.s.c
      const re = m.e.r
      const ce = m.e.c
      spanMap.set(`${rs},${cs}`, {
        rowSpan: re - rs + 1,
        colSpan: ce - cs + 1,
      })
      for (let r = rs; r <= re; r++) {
        for (let c = cs; c <= ce; c++) {
          if (r === rs && c === cs) continue
          covered.add(`${r},${c}`)
        }
      }
    }

    // 열 너비
    const colsInfo = ws['!cols'] || []
    const colWidths = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      colWidths.push(colPx(colsInfo[c]))
    }

    // 행 데이터
    const rows = []
    for (let r = range.s.r; r <= range.e.r; r++) {
      const rowInfo = (ws['!rows'] || [])[r]
      const cells = []
      for (let c = range.s.c; c <= range.e.c; c++) {
        const key = `${r},${c}`
        if (covered.has(key)) {
          cells.push(null) // 병합에 가려짐 → 렌더 안 함
          continue
        }
        const addr = XLSX.utils.encode_cell({ r, c })
        const cell = ws[addr]
        // 표시값: 서식 문자열(.w) 우선, 없으면 원시값(.v). 수식(.f)은 표시하지 않음.
        let text = ''
        let numeric = false
        if (cell) {
          if (cell.w != null) text = cell.w
          else if (cell.v != null) text = String(cell.v)
          numeric = cell.t === 'n'
        }
        const span = spanMap.get(key)
        cells.push({
          text,
          numeric,
          rowSpan: span?.rowSpan || 1,
          colSpan: span?.colSpan || 1,
        })
      }
      rows.push({ cells, hidden: rowInfo?.hidden === true })
    }

    return {
      name,
      startCol: range.s.c,
      startRow: range.s.r,
      colCount: range.e.c - range.s.c + 1,
      colWidths,
      rows,
    }
  })

  return { sheetNames: wb.SheetNames, sheets }
}

export default function XlsxView({ buffer }) {
  const model = useMemo(() => {
    try {
      return { data: buildModel(buffer), error: null }
    } catch (err) {
      console.error(err)
      return { data: null, error: err?.message || String(err) }
    }
  }, [buffer])

  const [activeSheet, setActiveSheet] = useState(0)

  if (model.error) {
    return (
      <div className="state-note error">
        스프레드시트를 열 수 없습니다: {model.error}
      </div>
    )
  }

  const { sheets } = model.data
  const sheet = sheets[Math.min(activeSheet, sheets.length - 1)]

  return (
    <div className="xlsx-wrap">
      <div className="xlsx-grid-scroll">
        <table className="xlsx-table">
          <colgroup>
            {/* 좌상단 행번호 열 */}
            <col style={{ width: 44 }} />
            {sheet.colWidths.map((w, i) => (
              <col
                key={i}
                style={w != null ? { width: w } : undefined}
              />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="corner" />
              {Array.from({ length: sheet.colCount }, (_, i) => (
                <th key={i} className="col-head">
                  {colLabel(sheet.startCol + i)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.rows.map((row, ri) => (
              <tr
                key={ri}
                style={row.hidden ? { display: 'none' } : undefined}
              >
                <th className="row-head">{sheet.startRow + ri + 1}</th>
                {row.cells.map((cell, ci) =>
                  cell === null ? null : (
                    <td
                      key={ci}
                      className={cell.numeric ? 'num' : undefined}
                      rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                      colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
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

      {/* 하단 시트 탭 */}
      <div className="sheet-tabs">
        {sheets.map((s, i) => (
          <button
            key={i}
            className={'sheet-tab' + (i === activeSheet ? ' active' : '')}
            onClick={() => setActiveSheet(i)}
            title={s.name}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
