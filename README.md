# 기찬엑셀 (GichanExcel)

가볍게 쓰는 Windows용 스프레드시트 앱입니다.  
엑셀·CSV 파일을 열고, 편집하고, 저장할 수 있습니다.

---

## 다운로드

[Releases](https://github.com/baggychani/GichanExcel/releases)에서 최신 버전을 받을 수 있습니다.

| 파일 | 설명 |
|------|------|
| **Setup (설치형)** | `GichanExcel_*_x64-setup.exe` — Windows에 설치합니다. xlsx / xls / csv / tsv / txt 파일을 더블클릭으로 바로 열 수 있습니다. |

---

## 주요 기능

- **파일 열기** — `.xlsx`, `.xls`, `.csv`, `.tsv`/`.txt` (xlsx는 SheetJS + OOXML 스타일 파서로 읽어, 기존보다 안정적으로 열립니다)
- **파일 저장** — `.xlsx`, `.csv`, `.tsv`/`.txt` (xlsx는 ExcelJS로 값·수식·서식·병합·행/열 크기를 저장합니다. 구형 `.xls` 저장은 지원하지 않습니다)
- **수식 계산** — `=SUM()`, `=IF()`, `=VLOOKUP()` 등 자주 쓰는 함수
- **서식** — 글꼴, 색, 테두리, 셀 병합, 조건부 서식
- **데이터 도구** — 정렬, 필터, 찾기·바꾸기, **텍스트를 열로 분할**
- **단축키** — `Ctrl+O` 열기, `Ctrl+S` 저장, `Ctrl+Shift+S` 다른 이름으로 저장
- **셀 줄바꿈** — `Alt+Enter` (엑셀과 동일), 긴 텍스트는 셀 안에서 자동 줄바꿈되어 다른 셀을 침범하지 않습니다
- **자동저장 + 종료 확인** — 실제 편집(값·스타일·행/열 크기 등)이 있을 때만 임시 저장되며, 저장하지 않은 채 닫으려 하면 저장 여부를 먼저 물어봅니다
- **시작 시 최대화** — 앱을 켜면 화면 크기에 맞게 최대화된 상태로 시작합니다. 최대화를 해제하면 1280×860(또는 화면에 맞는 크기)로 복원됩니다
- **업데이트 알림** — 앱 실행 시 GitHub Releases에 새 버전이 있으면 자동으로 알려줍니다

---

## 사용 방법

### 개발용 빠른 실행

소스 폴더에서 바로 실행하려면 **run.bat**을 더블클릭합니다.
Python 실행 연결이 되어 있다면 **app.py**도 같은 개발 모드 실행기로 사용할 수 있습니다.

> 배포받은 사용자는 소스 파일 대신 Releases의 Setup 설치 파일을 사용하는 것을 권장합니다.

### 설치 후 실행

1. Releases에서 **Setup** 파일을 다운로드합니다.
2. 설치 마법사를 따라 진행합니다.
3. 시작 메뉴 또는 바탕화면에서 **GichanExcel**을 실행합니다.

### 파일 열기

- 앱 상단 **열기** 버튼, 또는 `Ctrl+O`
- 탐색기에서 `.xlsx` 파일을 더블클릭 (설치형만 해당)

### 저장

- **저장** (`Ctrl+S`) — 현재 파일에 덮어쓰기
- **다른 이름으로 저장** (`Ctrl+Shift+S`) — 새 파일로 저장 (이미 저장된 파일이면 현재 파일명이 기본값)

### 텍스트를 열로 분할

1. 분할할 셀(또는 범위)을 선택합니다.
2. 상단 리본 **데이터** 탭 → **텍스트를 열로 분할**을 클릭합니다.
3. 구분자(탭, 쉼표, 공백 등)를 선택하고 **적용**합니다.

---

## 엑셀과의 호환성

| 항목 | 지원 여부 |
|------|-----------|
| xlsx / csv / tsv 열기·저장 | ✅ (xlsx 열기: SheetJS+OOXML, 저장: ExcelJS) |
| xls 열기 | ✅ (저장은 xlsx로 안내) |
| 기본 수식·서식 | ✅ |
| 조건부 서식, 필터, 정렬 | ✅ |
| 셀 메모 (앱 → xlsx 저장) | ✅ |
| VBA 매크로 | ❌ |
| 피벗 테이블 | 제한적 |
| 복잡한 차트 | 제한적 |

> 복잡한 엑셀 파일은 일부 서식이나 기능이 다르게 보일 수 있습니다. 중요한 파일은 **다른 이름으로 저장** 후 원본을 따로 보관해 두세요.

---

## 개발자용

### 필요 환경

- [Node.js](https://nodejs.org/) LTS
- [Rust](https://www.rust-lang.org/tools/install)
- Windows 10/11 (빌드·실행)

### 프로젝트 구조

```
GichanExcel/
├── .github/workflows/   # GitHub Actions (릴리스 빌드)
├── patches/             # Univer 등 서드파티 패치
├── public/
├── src/
│   ├── components/      # React UI (툴바, 아이콘 등)
│   ├── lib/             # 파일 I/O, 버전, 텍스트 분할
│   ├── plugins/         # Univer 플러그인 (텍스트 분할 메뉴)
│   ├── App.css
│   ├── main.tsx
│   └── setup-univer.ts  # Univer 초기화
├── src-tauri/           # Tauri (Rust) 데스크톱 셸
│   ├── capabilities/
│   ├── icons/
│   ├── src/
│   ├── Cargo.toml
│   └── tauri.conf.json
├── package.json
└── vite.config.ts
```

### 로컬에서 실행

```bash
npm install
npm run tauri dev
```

### 설치 파일 빌드

```bash
npm run tauri build
```

빌드 결과는 `src-tauri/target/release/bundle/nsis/` 에 생성됩니다.

- `GichanExcel_*_x64-setup.exe` — Setup 설치 파일

### 릴리스 배포

GitHub Actions **Release** 워크플로우에서 수동 실행하거나, 태그를 push하면 Setup 설치 파일이 자동 빌드됩니다.

**방법 1 — Actions에서 수동 실행 (권장)**

1. [Actions → Release](https://github.com/baggychani/GichanExcel/actions/workflows/release.yml) 로 이동
2. 우측 **Run workflow** 클릭
3. 버전 입력란을 **비워두면** `package.json`에 적힌 버전(현재 코드 기준)으로 릴리스합니다
4. 다른 버전으로 올릴 때만 직접 입력 (예: `1.1.1`) 후 실행

**방법 2 — 태그 push**

```bash
git tag v1.1.1
git push origin v1.1.1
```

---

## 기술 스택

- [Tauri 2](https://tauri.app/) — 데스크톱 앱
- [Univer](https://univer.ai/) — 스프레드시트 엔진
- [React](https://react.dev/) + [Vite](https://vitejs.dev/)

---

## 라이선스

MIT

Copyright (c) 2026 Bae Gichan

서드파티 고지는 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)를 참고하세요.
