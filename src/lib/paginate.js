// docx 자동 페이지 나눔 (워드처럼)
//
// docx-preview 는 명시적 페이지/구역 나누기에서만 페이지(section)를 나눈다. 워드가 아닌
// 프로그램(스크립트·라이브러리)이 만든 문서는 본문 전체가 거대한 section 하나로 나온다.
// 렌더 후 각 section 의 실제 크기(min-height = 용지 높이, padding = 여백)를 기준으로 블록
// (문단·제목·그림·표) 높이를 재서 넘치는 지점부터 같은 모양의 section 을 복제해 옮긴다.
// 표는 행 단위로 나누고 머리글 행(tblHeader / thead)은 다음 페이지에 반복한다.
//
// - 반드시 확대 100%(zoom 1)·화면에 보이는 상태에서 호출 (DocxView 가 보장)
// - 다시 실행(글꼴·그림 늦게 로드)하면 먼저 unpaginate 로 원래 구조로 되돌린 뒤 다시 나눈다.
//   (노드를 복제하지 않고 옮기기만 하므로 이미지 재로딩·무한 반복이 없다)
// - 생성 페이지: section[data-dv-page="cont"], 나뉜 표: table[data-dv-split], 이어진 표:
//   [data-dv-cont], 반복 머리글 행: tr[data-dv-rephead], 숨긴 끝 빈 문단: [data-dv-hidden]

const px = (v) => parseFloat(v) || 0
// 블록 아래 여백(margin-bottom). 페이지 마지막 블록의 아래 여백도 본문(article) 높이에 들어가
// 넘치면 section 이 용지보다 커진다(인쇄 시 빈 장) → 넘침 판정에 포함
const mb = (el) => px(getComputedStyle(el).marginBottom)
const EPS = 0.5 // 반올림 오차 허용(px)
let splitSeq = 0

// docx-preview 클래스 이름 규칙 (docx-preview 내부 escapeClassName 과 동일)
const escapeClassName = (s) =>
  s?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase()

// 본문 최상위 표마다 [머리글 행 수, 전체 행 수] — document.xml 을 직접 읽는다.
// (docx-preview 는 값 없는 <w:tblHeader/> 를 머리글로 인식하지 못함)
async function tableHeaderCounts(doc) {
  const part = doc.documentPart
  const xml = await part?._package?.load?.(part.path)
  if (!xml) return []
  const root = new DOMParser().parseFromString(xml, 'application/xml')
  const kids = (n, name) => [...n.children].filter((c) => c.localName === name)
  const on = (el) => !/^(0|false|off)$/i.test(el.getAttribute('w:val') ?? '')
  const out = []
  const walk = (n) => {
    for (const c of n.children) {
      if (c.localName === 'tbl') {
        const rows = kids(c, 'tr')
        let h = 0
        while (h < rows.length) {
          const pr = kids(rows[h], 'trPr')[0]
          const th = pr && kids(pr, 'tblHeader')[0]
          if (!th || !on(th)) break
          h++
        }
        out.push({ n: h, rows: rows.length })
      } else if (c.localName === 'sdt') {
        kids(c, 'sdtContent').forEach(walk) // docx-preview 처럼 내용 컨트롤은 펼친다
      }
    }
  }
  const body = root.getElementsByTagNameNS('*', 'body')[0]
  if (body) walk(body)
  return out
}

// 렌더 직후 1회: 표 머리글 행 수·'다음 문단과 함께' 스타일을 DOM 에 표시
export async function annotateDocx(wrapper, doc, className) {
  if (!wrapper || !doc) return
  try {
    // 1) 표 머리글 행(w:tblHeader) — 본문 최상위 표 순서 ↔ article > table 순서 대응
    const heads = await tableHeaderCounts(doc).catch(() => [])
    const tables = wrapper.querySelectorAll(':scope > section > article > table')
    tables.forEach((tb, i) => {
      const h = heads[i]
      // 행 수가 다르면 대응이 어긋난 것 → 표시 안 함(thead 가 있으면 그걸 사용)
      if (h && h.n > 0 && h.rows === tb.rows.length && h.n < tb.rows.length) {
        tb.dataset.dvHead = String(h.n)
      }
    })

    // 2) '다음 문단과 함께'(keepNext) 문단 스타일 — 제목이 페이지 끝에 홀로 남지 않게
    const styles = doc.stylesPart?.styles || []
    const byId = new Map(styles.map((s) => [s.id, s]))
    const keep = (s, depth = 0) => {
      if (!s || depth > 20) return false
      const k = s.paragraphProps?.keepNext
      if (k != null) return k
      if (/^(heading [1-9]|title|subtitle)$/i.test(s.name || '')) return true
      return keep(byId.get(s.basedOn), depth + 1)
    }
    const cls = styles
      .filter((s) => s.target === 'p' && s.id && keep(s))
      .map((s) => `${className}_${escapeClassName(s.id)}`)
    if (cls.length) {
      const sel = cls.map((c) => ':scope > section > article > p.' + CSS.escape(c)).join(',')
      wrapper.querySelectorAll(sel).forEach((p) => (p.dataset.dvKeep = '1'))
    }
  } catch (err) {
    console.warn('annotateDocx', err)
  }
}

// 자동 나눔을 되돌려 원래 section 구조로 복원
export function unpaginate(wrapper) {
  wrapper.querySelectorAll('[data-dv-hidden]').forEach((el) => {
    el.style.display = el.dataset.dvHidden === '-' ? '' : el.dataset.dvHidden
    delete el.dataset.dvHidden
  })
  const conts = [...wrapper.querySelectorAll(':scope > section[data-dv-page="cont"]')]
  for (const np of conts.reverse()) {
    const prev = np.previousElementSibling
    const prevArt = prev?.querySelector(':scope > article')
    const art = np.querySelector(':scope > article')
    if (!prevArt || !art) continue
    for (const child of [...art.children]) {
      const last = prevArt.lastElementChild
      if (
        child.tagName === 'TABLE' &&
        child.dataset.dvCont &&
        last?.tagName === 'TABLE' &&
        last.dataset.dvSplit === child.dataset.dvSplit
      ) {
        mergeTable(last, child)
      } else {
        prevArt.appendChild(child)
      }
    }
    // 각주 등 기타 요소는 원래 페이지(꼬리말 앞)로
    const prevFooter = prev.querySelector(':scope > footer')
    for (const c of [...np.children]) {
      if (!['HEADER', 'FOOTER', 'ARTICLE'].includes(c.tagName)) prev.insertBefore(c, prevFooter)
    }
    np.remove()
  }
}

function mergeTable(orig, cont) {
  for (const row of [...cont.rows]) {
    if (row.dataset.dvRephead) continue
    const p = row.parentNode
    let target = orig
    if (p !== cont) {
      const same = [...orig.children].filter((c) => c.tagName === p.tagName)
      target = same[same.length - 1]
      if (!target) {
        target = p.cloneNode(false)
        orig.appendChild(target)
      }
    }
    target.appendChild(row)
  }
}

const stripIds = (el) => {
  el.removeAttribute?.('id')
  el.querySelectorAll?.('[id]').forEach((n) => n.removeAttribute('id'))
  return el
}

// 다음 페이지(section) 생성: 같은 클래스·스타일, 머리말/꼬리말 복제, moved 블록 이동
function createPage(page, art, moved) {
  const np = page.cloneNode(false)
  np.dataset.dvPage = 'cont'
  const na = art.cloneNode(false)
  stripIds(na)
  const footer = page.querySelector(':scope > footer')
  for (const c of page.children) {
    if (c.tagName === 'HEADER') np.appendChild(stripIds(c.cloneNode(true)))
  }
  np.appendChild(na)
  for (const m of moved) na.appendChild(m)
  // 본문 뒤 각주 등(꼬리말 제외)은 마지막 페이지로 따라간다
  for (const c of [...page.children]) {
    if (c !== art && !['HEADER', 'FOOTER', 'ARTICLE'].includes(c.tagName)) np.appendChild(c)
  }
  if (footer) np.appendChild(stripIds(footer.cloneNode(true)))
  page.after(np)
  return np
}

const isEmptyBlock = (el) =>
  el.tagName === 'P' &&
  !el.textContent.trim() &&
  !el.querySelector('img,svg,table,canvas,object,iframe')

const hide = (el) => {
  el.dataset.dvHidden = el.style.display || '-'
  el.style.display = 'none'
}

// 표 정보: 행, 머리글 행 수, 나눌 수 없는 행(세로 병합 rowSpan 이 가로지르는 지점)
function tableInfo(table) {
  const rows = [...table.rows]
  let headN = +table.dataset.dvHead || (table.tHead ? table.tHead.rows.length : 0)
  if (headN >= rows.length) headN = 0
  const blocked = new Array(rows.length).fill(false)
  rows.forEach((r, j) => {
    for (const c of r.cells) {
      if (c.rowSpan > 1 && c.style.display !== 'none') {
        for (let k = j + 1; k < Math.min(rows.length, j + c.rowSpan); k++) blocked[k] = true
      }
    }
  })
  return { rows, headN, blocked }
}

// 표를 나눌 행 번호 k 를 고른다(행 k 부터 다음 페이지). 못 나누면 -1.
// over(j) = 행 j 가 페이지를 넘치는가, first = 표(조각)가 페이지 맨 위인가,
// from = 이 조각의 첫 행(이미 앞 페이지로 간 행 다음)
function chooseRowCut(info, over, first, from) {
  const { rows, headN, blocked } = info
  if (rows.length < 2) return -1
  let k = -1
  for (let j = from; j < rows.length; j++) {
    if (over(j)) {
      k = j
      break
    }
  }
  if (k < 0) return -1
  const minK = from > 0 ? from + 1 : headN + 1 // 이 페이지에 본문 행이 최소 1개
  while (k > minK && blocked[k]) k--
  if (k < minK) {
    // 본문 행이 한 줄도 안 들어감: 페이지 중간이면 표 전체를 다음 페이지로,
    // 페이지 맨 위라면 최소 1행은 남기고(넘침 허용) 나눈다
    if (!first) return -1
    k = minK
  }
  while (k < rows.length && blocked[k]) k++
  return k < rows.length ? k : -1
}

// 표를 행 k 에서 나눠 다음 페이지로 갈 새 표(머리글 행 반복)를 돌려준다
function splitTableAt(table, k, headN) {
  const rows = [...table.rows]
  if (!table.dataset.dvSplit) table.dataset.dvSplit = String(++splitSeq)
  const nt = table.cloneNode(false)
  stripIds(nt)
  nt.dataset.dvCont = '1'
  for (const c of table.children) {
    if (c.tagName === 'COLGROUP') nt.appendChild(c.cloneNode(true))
  }
  const map = new Map()
  const holder = (p) => {
    if (p === table) return nt
    let h = map.get(p)
    if (!h) {
      h = p.cloneNode(false)
      nt.appendChild(h)
      map.set(p, h)
    }
    return h
  }
  for (let i = 0; i < headN; i++) {
    const h = stripIds(rows[i].cloneNode(true))
    h.dataset.dvRephead = '1'
    holder(rows[i].parentNode).appendChild(h)
  }
  for (let i = k; i < rows.length; i++) holder(rows[i].parentNode).appendChild(rows[i])
  return nt
}

// [1단계] 한 번의 레이아웃 측정으로 구역 전체의 나눌 지점을 계획한 뒤 한꺼번에 나눈다.
// (페이지마다 다시 측정하면 30쪽 문서에서 수십 번 전체 레이아웃 → 느림)
// 새 페이지에서 블록 위치 = 원래 위치 - shift (shift = 새 페이지 첫 블록이 본문 맨 위로
// 올라가는 거리). 근사이므로 이후 2단계(splitPage)가 실제로 넘치는 페이지만 보정한다.
function planAndSplit(sec, art, limit) {
  const secTop = sec.getBoundingClientRect().top
  const artTop = art.getBoundingClientRect().top - secTop
  const kids = [...art.children]
  const n = kids.length
  const R = kids.map((el) => el.getBoundingClientRect())
  const mt = (el) => px(getComputedStyle(el).marginTop)
  const B = kids.map((el, i) => (R[i].height ? R[i].bottom + mb(el) : 0))
  const lim = limit - 2 // 근사 오차 여유
  const cuts = [] // { i, k?, headN? } — 블록 i(표면 행 k)부터 새 페이지
  let hideFrom = -1
  let shift = 0
  let pageStart = 0 // 현재 페이지 첫 블록
  let rowFrom = 0 // 현재 페이지 첫 블록이 이어진 표일 때 그 조각의 첫 행
  let tinfo = null
  let i = 0
  while (i < n) {
    if (R[i].height === 0 || B[i] - secTop - shift <= lim) {
      i++
      rowFrom = 0
      continue
    }
    // 넘친 것이 전부 끝의 빈 문단뿐이면 빈 페이지를 만들지 않고 숨긴다
    if (i > pageStart && kids.slice(i).every(isEmptyBlock)) {
      hideFrom = i
      break
    }
    const el = kids[i]
    if (el.tagName === 'TABLE') {
      const info = (tinfo = tinfo?.table === el ? tinfo : { table: el, ...tableInfo(el) })
      const rr = info.rr || (info.rr = info.rows.map((r) => r.getBoundingClientRect()))
      const tmb = B[i] - R[i].bottom // 표 아래 여백
      const k = chooseRowCut(
        info,
        (j) => rr[j].bottom + tmb - secTop - shift > lim,
        i === pageStart,
        i === pageStart ? rowFrom : 0
      )
      if (k > 0) {
        cuts.push({ i, k, headN: info.headN })
        const headH = info.headN ? rr[info.headN - 1].bottom - rr[0].top : 0
        // 새 페이지: 표가 맨 위 → 첫 행 위치 = 본문 위 + 표 margin + (표 위~첫 행) + 머리글
        shift = rr[k].top - secTop - (artTop + mt(el) + (rr[0].top - R[i].top) + headH)
        pageStart = i
        rowFrom = k
        continue // 같은 표의 나머지 행을 새 페이지 기준으로 다시 확인
      }
    }
    let cut = i
    if (cut === pageStart) {
      cut = i + 1 // 한 블록이 페이지보다 큼 → 그대로 넘치게 두고 다음 블록부터 새 페이지
    } else {
      // '다음 문단과 함께'(제목 등) 문단은 넘친 블록과 같이 다음 페이지로
      while (cut > pageStart && kids[cut - 1].dataset.dvKeep) cut--
      if (cut === pageStart) cut = i
    }
    if (cut >= n) break
    cuts.push({ i: cut })
    shift = R[cut].top - secTop - (artTop + mt(kids[cut]))
    pageStart = i = cut
    rowFrom = 0
  }

  // 측정 끝 → DOM 변경(뒤에서부터: 앞쪽 인덱스·행 번호가 그대로 유효)
  if (hideFrom >= 0) kids.slice(hideFrom).forEach(hide)
  for (let c = cuts.length - 1; c >= 0; c--) {
    const { i: ci, k, headN } = cuts[c]
    const rest = kids.slice(ci + 1).filter((el) => el.parentNode === art)
    const moved = k != null ? [splitTableAt(kids[ci], k, headN), ...rest] : [kids[ci], ...rest]
    createPage(sec, art, moved)
  }
}

// [2단계] 한 페이지에서 실제로 넘치는 지점을 찾아 다음 페이지를 만든다. 없으면 null
function splitPage(page, limit) {
  const art = page.querySelector(':scope > article')
  if (!art) return null
  const top = page.getBoundingClientRect().top
  const kids = [...art.children]
  const i = kids.findIndex((el) => {
    const r = el.getBoundingClientRect()
    return r.height > 0 && r.bottom + mb(el) - top > limit + EPS
  })
  if (i < 0) return null

  const rest = kids.slice(i)
  if (i > 0 && rest.every(isEmptyBlock)) {
    rest.forEach(hide)
    return null
  }

  const el = kids[i]
  if (el.tagName === 'TABLE') {
    const info = tableInfo(el)
    // 이어진 표 조각은 반복 머리글 행 다음이 첫 본문 행
    const from = el.dataset.dvCont ? info.headN : 0
    const tmb = mb(el)
    const k = chooseRowCut(
      info,
      (j) => info.rows[j].getBoundingClientRect().bottom + tmb - top > limit + EPS,
      i === 0,
      from
    )
    if (k > 0) {
      return createPage(page, art, [splitTableAt(el, k, info.headN), ...kids.slice(i + 1)])
    }
  }
  let cut = i
  if (cut === 0) {
    cut = 1
  } else {
    while (cut > 0 && kids[cut - 1].dataset.dvKeep) cut--
    if (cut === 0) cut = i
  }
  const moved = kids.slice(cut)
  if (!moved.length) return null
  return createPage(page, art, moved)
}

function paginateSection(sec) {
  const cs = getComputedStyle(sec)
  const pageH = px(cs.minHeight)
  if (!pageH) return
  const arts = [...sec.children].filter((c) => c.tagName === 'ARTICLE')
  // 다단(column) 구역은 블록 높이로 나눌 수 없어 그대로 둔다
  if (arts.length !== 1 || arts[0].style.columnCount) return
  // 꼬리말이 여백보다 크면 그만큼 본문 영역이 줄어든다
  const footer = sec.querySelector(':scope > footer')
  const extraFoot = footer
    ? Math.max(0, footer.offsetHeight + px(getComputedStyle(footer).marginBottom))
    : 0
  const limit = pageH - px(cs.paddingBottom) - extraFoot // section 위쪽 기준 본문 하한
  planAndSplit(sec, arts[0], limit)
  // 보정: 실제로 넘치는 페이지만 다시 나눈다(대개 없음 → 레이아웃 1회)
  let page = sec
  for (let guard = 0; page && guard < 5000; guard++) {
    splitPage(page, limit)
    page = page.nextElementSibling
    if (page?.dataset.dvPage !== 'cont') break
  }
}

// 전체 문서 자동 페이지 나눔 (다시 호출해도 안전)
export function paginate(wrapper) {
  if (!wrapper) return
  unpaginate(wrapper)
  for (const sec of [...wrapper.querySelectorAll(':scope > section')]) paginateSection(sec)
}
