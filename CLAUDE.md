# CLAUDE.md — 프로젝트 가이드 (AI/개발자용)

로컬 `.docx` / `.xlsx` **뷰어 + 간이 편집기**. 업로드 없이 브라우저/데스크톱
앱에서 열람·편집. **폐쇄망 대응**(외부 CDN·네트워크 없이 동작).

## 구성 요소 (3가지 배포 형태)

1. **개발 서버** — `npm run dev` (Vite + React)
2. **단일 HTML** — `release/문서뷰어.html` (더블클릭 → 브라우저, 오프라인)
3. **윈도우 설치 앱** — Electron → nsis 설치 파일 (파일연결·아이콘 자동 등록)

## 디렉터리

```
src/
  App.jsx              파일 상태/드래그앤드롭/탭. 뷰어는 React.lazy 로 지연 로딩
  views/
    DocxView.jsx       보기=docx-preview, 편집=mammoth(지연), 저장=html-to-docx
    XlsxView.jsx       SheetJS 파싱/렌더, 셀 편집, 저장
  lib/save.js          저장(Electron 대화상자 / 브라우저 다운로드), docx export
electron/
  main.cjs             메인 프로세스(소스). 파일연결 열기 + 저장 IPC
  main.build.cjs        ← esbuild 번들 산출(gitignore). package.json "main"
  preload.cjs          contextBridge 브리지
  afterPack.cjs        빌드 후 언어팩 제거(용량)
build/icon.ico         앱/파일 아이콘
.github/workflows/build-exe.yml   윈도우에서 설치 파일 자동 빌드
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

1. `package.json` 의 `version` 을 올린다 (예: 1.1.0 → 1.2.0)
2. 커밋 & `claude/local-doc-viewer-5o6qvl` 브랜치에 push
3. GitHub **Actions → Build Windows EXE** 가 자동 실행
4. 완료 후 **Artifacts → 문서뷰어-Setup** 다운로드 → `문서뷰어-Setup-<버전>.exe`

> 버전은 설치 파일명과 앱 좌측 상단(`__APP_VERSION__`, vite define)에 표시된다.

---

## ⚠️ 윈도우 패키징 함정 (하드코어 교훈 — 반드시 참고)

이 프로젝트를 만들며 실제로 부딪힌 문제와 해법. 새로 손대기 전에 읽을 것.

- **아이콘 ≠ 파일연결**: "연결 프로그램"은 여는 앱만 바꾸고 파일 아이콘은 안
  바뀐다. 아이콘까지 바꾸려면 **설치형(nsis) + `fileAssociations`** 로 정식
  등록해야 한다. 포터블 exe 로는 최신 윈도우에서 아이콘 변경이 사실상 불가.
- **포터블 exe 는 느리다**: `portable` 타깃은 실행마다 임시폴더에 자기 압축을
  푼다. **`nsis`(설치형) 또는 `zip`(폴더형)** 이 훨씬 빠르다. → nsis 채택.
- **SmartScreen**: 서명 안 된 exe 는 최초 실행 시 경고. 코드서명 인증서(유료)
  없으면 "추가 정보 → 실행" 또는 파일 속성 → 차단 해제(=MOTW 제거)로 우회.
- **electron-builder CI 실패 `GH_TOKEN not set`**: nsis 는 CI 에서 GitHub
  Releases 자동 배포를 시도한다. 배포 안 쓰면 **`--publish never`** 필수.
- **메인 프로세스의 런타임 의존성**: 렌더러는 Vite 가 번들하지만 `electron/
  main.cjs` 의 `require('html-to-docx')` 는 런타임에 node_modules 를 찾는다.
  패키징 누락 위험을 없애려 **esbuild 로 main 을 번들**(`build:main`)해서
  라이브러리를 인라인 → 앱은 dist + main.build.cjs + preload.cjs 로 자립.
- **html-docx-js 는 번들 불가**: `with` 문 때문에 Rollup/Vite 에서 실패.
  docx 생성은 **html-to-docx**(네이티브 WordML) 를 메인 프로세스에서 사용.
- **용량**: Electron 은 최소 ~70MB(브라우저 엔진 포함). `compression: maximum`
  + `afterPack.cjs` 로 미사용 언어팩 제거(한/영만)로 절감. 더 줄이려면 Tauri
  (시스템 웹뷰, ~10MB) 재작성이 필요 — 큰 작업.
- **속도**: 무거운 파서(xlsx/mammoth/docx-preview)는 `React.lazy`/동적 import
  로 필요할 때만 로드. 창은 `ready-to-show` 후 표시 + `backgroundColor` 로
  흰 화면 깜빡임 제거.
- **한글 productName/파일명**: electron-builder 에서 정상 동작(검증됨).

## xlsx 렌더/편집 (ExcelJS)

- **XlsxView 는 ExcelJS 로 읽는다**(SheetJS 아님). 셀 채우기색·글꼴(굵게/색/
  크기)·테두리·정렬·병합·열너비를 모두 읽어 **엑셀과 비슷하게 렌더**한다.
- 테마색은 `THEME` 팔레트 + tint 로 근사 변환(`colorToCss`).
- **편집 저장 시 스타일이 보존된다**(원본 워크북을 유지한 채 값만 반영 →
  `wb.xlsx.writeBuffer()`). SheetJS 무료판의 "저장 시 서식 손실" 문제 없음.

## 편집 기능의 한계 (정직하게)

- **xlsx**: 값·수식·병합·열너비·색·서식 유지. 셀 값 편집 위주(행/열 삽입·수식
  편집기는 없음).
- **docx**: mammoth 로 단순 HTML 편집(+B/I/U 서식) → html-to-docx 로 저장. 기본
  요소(문단·굵게·표) 위주. 원본의 정교한 서식은 일부 손실. **워드 저장은 앱
  (.exe) 전용**(브라우저는 편집만). 최종 서식 중요 문서는 워드에서 마무리.
