import React from 'react'

// 툴바용 작은 선 아이콘 (외부 아이콘 폰트 없이 인라인 SVG — 폐쇄망 대응)
// 정렬: 왼쪽/가운데/오른쪽 — 워드·엑셀의 정렬 아이콘 모양
const ALIGN_LINES = {
  left: [[2, 14], [2, 10], [2, 14], [2, 9]],
  center: [[2, 14], [4, 12], [2, 14], [4.5, 11.5]],
  right: [[2, 14], [6, 14], [2, 14], [7, 14]],
}

export function AlignIcon({ dir = 'left' }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      {ALIGN_LINES[dir].map(([x1, x2], i) => (
        <line
          key={i}
          x1={x1}
          x2={x2}
          y1={3 + i * 3.4}
          y2={3 + i * 3.4}
          stroke="currentColor"
          strokeWidth="1.6"
          strokeLinecap="round"
        />
      ))}
    </svg>
  )
}
