# 기찬엑셀 (GichanExcel)

가볍게 쓰는 Windows용 스프레드시트 앱입니다.  
엑셀·CSV 파일을 열고, 편집하고, 저장할 수 있습니다.

---

## 다운로드

[Releases](https://github.com/baggychani/GichanExcel/releases)에서 최신 버전을 받을 수 있습니다.

| 파일 | 설명 |
|------|------|
| **Setup (설치형)** | `기찬엑셀_*_x64-setup.exe` — Windows에 설치합니다. xlsx / xls / csv 파일을 더블클릭으로 바로 열 수 있습니다. |

---

## 주요 기능

- **파일 열기 / 저장** — `.xlsx`, `.xls`, `.csv` 지원
- **수식 계산** — `=SUM()`, `=IF()`, `=VLOOKUP()` 등 자주 쓰는 함수
- **서식** — 글꼴, 색, 테두리, 셀 병합, 조건부 서식
- **데이터 도구** — 정렬, 필터, 찾기·바꾸기, **텍스트를 열로 분할**
- **단축키** — `Ctrl+O` 열기, `Ctrl+S` 저장, `Ctrl+Shift+S` 다른 이름으로 저장
- **셀 줄바꿈** — `Alt+Enter` (엑셀과 동일)

---

## 사용 방법

### 설치 후 실행

1. Releases에서 **Setup** 파일을 다운로드합니다.
2. 설치 마법사를 따라 진행합니다.
3. 시작 메뉴 또는 바탕화면에서 **기찬엑셀**을 실행합니다.

### 파일 열기

- 앱 상단 **열기** 버튼, 또는 `Ctrl+O`
- 탐색기에서 `.xlsx` 파일을 더블클릭 (설치형만 해당)

### 저장

- **저장** (`Ctrl+S`) — 현재 파일에 덮어쓰기
- **다른 이름으로 저장** (`Ctrl+Shift+S`) — 새 파일로 저장

### 텍스트를 열로 분할

1. 분할할 셀(또는 범위)을 선택합니다.
2. 상단 리본 **데이터** 탭 → **텍스트를 열로 분할**을 클릭합니다.
3. 구분자(탭, 쉼표, 공백 등)를 선택하고 **적용**합니다.

---

## 엑셀과의 호환성

| 항목 | 지원 여부 |
|------|-----------|
| xlsx / xls / csv 열기·저장 | ✅ |
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

- `기찬엑셀_*_x64-setup.exe` — Setup 설치 파일

### 릴리스 배포

GitHub Actions **Release** 워크플로우에서 수동 실행하거나, 태그를 push하면 Setup 설치 파일이 자동 빌드됩니다.

**방법 1 — Actions에서 수동 실행 (권장)**

1. [Actions → Release](https://github.com/baggychani/GichanExcel/actions/workflows/release.yml) 로 이동
2. 우측 **Run workflow** 클릭
3. 버전 입력 (예: `1.0.1`) 후 실행

**방법 2 — 태그 push**

```bash
git tag v1.0.1
git push origin v1.0.1
```

---

## 기술 스택

- [Tauri 2](https://tauri.app/) — 데스크톱 앱
- [Univer](https://univer.ai/) — 스프레드시트 엔진
- [React](https://react.dev/) + [Vite](https://vitejs.dev/)

---

## 라이선스

MIT
