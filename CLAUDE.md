# CLAUDE.md — 프로젝트 가이드 (AI/개발자용)

로컬 문서 **뷰어 + 간이 편집기**. 업로드 없이 브라우저/데스크톱 앱에서 열람·편집.
**폐쇄망 대응**(외부 CDN·네트워크 없이 동작).

지원 형식: `.docx` `.xlsx` `.xls` `.csv` `.txt` `.log` `.md` `.ini` `.bat`

## 구성 요소 (3가지 배포 형태)

1. **개발 서버** — `npm run dev` (Vite + React)
2. **단일 HTML** — `release/문서뷰어.html` (더블클릭 → 브라우저, 오프라인)
3. **윈도우 설치 앱** — Electron → nsis 설치 파일 (파일연결·아이콘 자동 등록)

## 디렉터리

```
src/
  App.jsx              파일 상태/드래그앤드롭/탭/확대축소/찾기/우클릭메뉴/최초실행안내
  views/
    DocxView.jsx       보기=docx-preview, 편집=mammoth(지연)+execCommand, 저장=html-to-docx
                       A4 페이지 자동맞춤(onAutoFit)
    XlsxView.jsx       ExcelJS 로 렌더/편집(색·서식 보존). .xls/.csv 는 SheetJS(지연)
                       로 xlsx 변환 후 처리. 범위선택/행열삽입삭제/셀서식
    TextView.jsx       txt/log/ini/bat 텍스트 + md 마크다운(marked, 지연). 편집·저장
  lib/save.js          저장(Electron 대화상자 / 브라우저 다운로드), exportDocxSave
electron/
  main.cjs             메인 프로세스(소스). 파일연결 열기 + 저장/바로가기/기본앱 IPC
  main.build.cjs        ← esbuild 번들 산출(gitignore). package.json "main"
  preload.cjs          contextBridge 브리지
  afterPack.cjs        빌드 후 언어팩 제거(용량)
build/icon.ico         앱/파일 아이콘
.github/workflows/build-exe.yml   윈도우 설치 파일 자동 빌드 + Release 자동 생성
```

## 자주 쓰는 명령

```bash
npm install
npm run dev            # 개발 서버
npm run build          # dist/ (렌더러)
npm run build:single   # release/ 단일 HTML (SINGLE=1)
npm run dist:win       # 윈도우 설치 파일(release-exe/) — 윈도우에서 실행
```

## 릴리스 방법 (버전 올리기)

1. `package.json` 의 `version` 을 올린다 (예: 1.6.0 → 1.7.0)
2. 커밋 & `claude/local-doc-viewer-5o6qvl` 브랜치에 push
3. GitHub **Actions → Build Windows EXE** 가 자동 실행
4. 완료 후 **Releases** 에 `v<버전>` 자동 생성 + 설치 파일 첨부(상시 링크).
   동료 공유용 고정 링크: `github.com/<owner>/viewer/releases/latest`
   (Actions 의 Artifacts 로도 받을 수 있음 — 90일 보관)

> 버전은 설치 파일명과 앱 좌측 상단(`__APP_VERSION__`, vite define)에 표시된다.

---

## ⚠️ 윈도우 패키징·배포 함정 (하드코어 교훈 — 반드시 참고)

이 프로젝트를 만들며 실제로 부딪힌 문제와 해법. 새로 손대기 전에 읽을 것.

### 빌드/패키징
- **아이콘 ≠ 파일연결**: "연결 프로그램"은 여는 앱만 바꾸고 파일 아이콘은 안
  바뀐다. 아이콘까지 바꾸려면 **설치형(nsis) + `fileAssociations`** 로 정식
  등록해야 한다. 포터블 exe 로는 최신 윈도우에서 아이콘 변경이 사실상 불가.
- **포터블 exe 는 느리다**: `portable` 타깃은 실행마다 임시폴더에 자기 압축을
  푼다. **`nsis`(설치형)** 채택. (zip=폴더형도 빠르지만 아이콘 등록 안 됨)
- **메인 프로세스 런타임 의존성**: 렌더러는 Vite 가 번들하지만 `electron/main.cjs`
  의 `require('html-to-docx')` 는 런타임에 node_modules 를 찾는다. 누락 위험을
  없애려 **esbuild 로 main 을 번들**(`build:main`)해 라이브러리를 인라인 → 앱은
  dist + main.build.cjs + preload.cjs 로 자립(런타임 node_modules 불필요).
- **html-docx-js 는 번들 불가**: `with` 문 때문에 Rollup/Vite 에서 실패. docx
  생성은 **html-to-docx**(네이티브 WordML) 를 메인 프로세스에서 사용.
- **용량**: Electron 최소 ~70MB. `compression: maximum` + `afterPack.cjs` 로
  미사용 언어팩 제거(한/영만) → 설치파일 ~76MB. 더 줄이려면 Tauri 재작성(큰 작업).
- **한글 productName/파일명**: electron-builder 에서 정상 동작(검증됨).

### 설치 UX (nsis)
- **assisted 설치(oneClick:false)** 를 쓴다. oneClick(무설치UI)은 어디 깔리는지
  안 보이고 재실행 시 조용히 재설치돼 혼란. assisted 는 설치 위치 표시/선택 +
  완료화면 실행 + **재설치 시 덮어쓰기 업데이트**(같은 appId → 중복 아님).
- **`perMachine: false`** → 관리자 권한 없이 사용자 계정 설치(폐쇄망 적합).
  설치 위치: `%LOCALAPPDATA%\Programs\문서뷰어`.

### 윈도우가 "막아둔" 것 — 자동화 불가, 앱에서 우회
- **파일연결 ≠ 기본 프로그램**: 설치가 fileAssociations 를 등록해도 Win10/11 은
  **기본앱(default)을 설치 프로그램이 강제 변경 못 하게** 막는다(UserChoice 해시
  보호). 이미 Word/Excel 이 기본인 확장자는 **사용자가 1회 직접 지정**해야 한다.
  → 앱 최초 실행 안내창에서 `ms-settings:defaultapps` 를 열어 유도(`open-default-apps`).
- **작업 표시줄 고정 불가**: MS 가 프로그래밍 방식 pin 을 차단. 설치/코드로 자동
  고정 못 함. → 바탕화면 바로가기(`shell.writeShortcutLink`) 후 "우클릭 → 작업
  표시줄에 고정" 안내가 최선.
- **SmartScreen**: 서명 안 된 exe/설치파일은 최초 1회 "PC 보호" 경고. **근본
  해결은 코드 서명 인증서(유료)뿐.** 없으면 "추가 정보 → 실행" 또는 파일 속성 →
  차단 해제(MOTW 제거). 설치 후 시작메뉴 실행은 경고 없음(설치파일에만 1회).
- **.bat 등 실행형은 파일연결 등록 금지**: 뷰어에 연결하면 더블클릭 시 스크립트
  실행이 막혀 오히려 불편. fileAssociations 는 오피스 형식(docx/xlsx/xls/csv)만.
  텍스트/실행형은 드롭·"연결 프로그램"으로만 열게 둔다.

### CI / Release
- **electron-builder CI 실패 `GH_TOKEN not set`**: nsis 는 CI 에서 자동 배포를
  시도한다. **`--publish never`** 로 끈다(dist:win 에 포함).
- **Release 자동 생성**은 electron-builder publish 대신 **softprops/action-gh-release@v2**
  로 별도 처리. 워크플로에 **`permissions: contents: write`** 필요. package.json
  version 으로 `v<버전>` 태그 생성/갱신(같은 버전 재push 시 자산 덮어씀).
- 산출물은 아티팩트(90일) + Release(상시) 둘 다 업로드.

### 렌더러 성능
- 무거운 파서(exceljs/mammoth/docx-preview/xlsx)는 `React.lazy`/동적 import 로
  필요할 때만 로드(초기 셸 ~150KB 유지). 창은 `ready-to-show` 후 표시 +
  `backgroundColor` 로 흰 화면 깜빡임 제거.

---

## 포맷별 처리 요약

- **xlsx** — ExcelJS 로 읽어 채우기색·글꼴·테두리·정렬·병합·열너비를 렌더. 테마색은
  `THEME` 팔레트+tint 근사(`colorToCss`). 편집 저장 시 **스타일 보존**(원본 wb 유지
  → `wb.xlsx.writeBuffer()`). SheetJS 무료판의 서식 손실 문제 없음.
- **xls(구형)/csv** — ExcelJS 는 xls 를 못 읽는다. **SheetJS(지연 로딩)** 로 읽어
  xlsx 버퍼로 변환 후 ExcelJS 파이프라인 재사용. CSV 는 UTF-8(BOM) 우선, 깨지면
  **EUC-KR 폴백**(한글 엑셀 CSV 대응). 편집 저장은 `.xlsx` 로(서식 보존).
- **docx** — 보기=docx-preview(충실). 편집=mammoth 로 단순 HTML 화 후 contentEditable
  (+execCommand 서식). 저장=html-to-docx(네이티브). **워드 저장은 앱(.exe) 전용**
  (브라우저는 편집만). A4 페이지는 열 때 창 너비에 자동 맞춤.
- **txt/log/ini/bat** — 모노스페이스 텍스트 뷰(+줄바꿈 토글), 편집·저장(UTF-8).
- **md** — marked(지연)로 마크다운 렌더. 원문 편집 가능. innerHTML 최소 방어(sanitize).

## 편집 기능의 한계 (정직하게)

- **xlsx**: 값·수식·병합·열너비·색·서식 유지. 셀 서식(굵게/취소선/크기/색/채우기/
  정렬)·범위선택(Shift+클릭)·행열 삽입삭제 지원. 수식 편집기는 없음.
- **docx**: 기본 요소(문단·굵게·표) 위주. 취소선은 html-to-docx 미지원. 원본의
  정교한 서식은 일부 손실 → 최종 서식 중요 문서는 워드에서 마무리.
- **공통**: 원본 덮어쓰지 않고 `이름-편집.*` 새 파일로 저장.

## 안 되는 것 / 큰 작업

- **ppt/pptx**: 슬라이드·도형·애니메이션 재현이 어려워 브라우저 렌더 품질 낮음
  (.ppt 구형은 라이브러리 사실상 없음). 필요하면 **PDF 로 내보내 PDF 뷰어**로
  보는 방식을 권장(정확·안정).
- **용량 대폭 축소**: Tauri(시스템 웹뷰) 재작성 필요 — 큰 작업.
