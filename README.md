# 로컬 문서 뷰어 (Local Doc Viewer)

`.docx` / `.xlsx` 파일을 **업로드 없이 브라우저에서** 열람하는 로컬 웹 뷰어입니다.
모든 파싱·렌더링이 브라우저 안에서만 일어나며, 파일이 외부 서버로 전송되지 않습니다.

**폐쇄망 대응**: 외부 CDN을 전혀 사용하지 않고 모든 의존성을 번들에 포함합니다.
빌드된 `dist/` 폴더만 있으면 오프라인 환경(사내망/에어갭)에서 그대로 동작합니다.

## 기능

- **파일 입력**
  - 드래그 앤 드롭 + 파일 선택 버튼
  - `.docx` `.xlsx` 확장자 필터 (그 외 파일은 자동 제외)
  - 여러 파일을 상단 탭 + 좌측 리스트로 동시에 열기
- **DOCX 렌더링** — [`docx-preview`](https://www.npmjs.com/package/docx-preview) 사용
  - 스타일 / 표 / 이미지 보존 (이미지는 data URL로 인라인 → 외부 요청 없음)
  - `mammoth`는 서식 손실이 커서 서식 보존 목적에는 `docx-preview`를 채택
- **XLSX 렌더링** — [`SheetJS(xlsx)`](https://www.npmjs.com/package/xlsx) + 자체 HTML table
  - 시트가 여러 개면 하단 탭으로 전환
  - 셀 병합(`!merges`) → `rowSpan` / `colSpan` 반영
  - 열 너비(`!cols`) → `<colgroup>` 픽셀 너비로 반영
  - **수식은 결과값으로 표시** — `cellFormula:false`로 읽고 서식값(`.w`)/원시값(`.v`)만 렌더 (수식 문자열 `=B3+B4` 노출 방지)
- **UI** — 좌: 파일 리스트 / 우: 렌더 영역, 상단 파일 탭

## 개발

```bash
npm install     # 의존성 설치 (최초 1회, 네트워크 필요)
npm run dev     # 개발 서버 (http://localhost:5173)
```

## 빌드 & 오프라인 배포

```bash
npm run build   # dist/ 생성 (모든 의존성 번들 포함, CDN 참조 없음)
npm run preview # 빌드 결과 로컬 확인
```

빌드 후 생성된 `dist/` 폴더를 폐쇄망 환경으로 복사하면 정적 파일 서버(또는 브라우저에서 직접)로 열 수 있습니다.

## 기술 스택

| 영역 | 라이브러리 |
|------|-----------|
| 앱 | Vite + React 18 |
| DOCX | docx-preview |
| XLSX | SheetJS (xlsx) + 자체 HTML table |

## 구조

```
src/
  main.jsx          진입점
  App.jsx           파일 상태 / 드래그앤드롭 / 파일 리스트 / 탭
  styles.css        전역 스타일
  views/
    DocxView.jsx    docx-preview 렌더
    XlsxView.jsx    SheetJS 파싱 + 병합/열너비/수식결과 렌더
```
