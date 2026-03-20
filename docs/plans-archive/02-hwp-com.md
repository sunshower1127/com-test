# Phase 2: HWP COM 자동화 (Rust)

## 목적

Phase 1에서 확립한 COM 패턴을 HWP에 적용한다.
HWP 고유의 보안 모듈 등록, 버전 파편화 대응을 추가로 해결한다.

## 사전 조건

- Windows 11
- 한컴오피스 설치됨 (2020/2022/NEO 등)
- Phase 1 `com-core` 완성

## HWP 고유 과제

### 1. 보안 모듈 등록 (최대 장벽)

HWP COM은 파일 접근 전 보안 DLL 등록이 필수.

```
RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
```

- DLL은 한컴오피스 설치 폴더에 존재
- 설치 경로는 레지스트리에서 동적 탐색 필요
- 버전마다 레지스트리 키가 다름 → fallback 탐색

### 2. 레지스트리 탐색 경로 (알려진 키들)

```
HKLM\SOFTWARE\HNC\HwpAutomation\Versions
HKLM\SOFTWARE\Hnc\Hwp\CurrentVersion
HKLM\SOFTWARE\WOW6432Node\HNC\...
```

### 3. 버전 파편화

| 버전 | COM ProgID | 비고 |
|------|-----------|------|
| 한/글 2018+ | HWPFrame.HwpObject | 일반적 |
| 한컴오피스 NEO | HWPFrame.HwpObject | 동일하지만 DLL 경로 다름 |

## HWP COM 객체 계층

```
HWPFrame.HwpObject
├── .RegisterModule(...)          // 보안 모듈 등록
├── .Visible                      // bool
├── .Open(path)                   // 문서 열기
├── .SaveAs(path, format)         // 저장
├── .GetTextFile("TEXT", "")      // 전체 텍스트 추출
├── .HAction
│   ├── .GetDefault(action, paramset)
│   └── .Execute(action, paramset)
├── .HParameterSet
│   └── .HInsertText
│       └── .Text                 // 삽입할 텍스트
├── .XHwpDocuments                // 열린 문서 목록
└── .Quit()
```

## 구현 단계

### Step 1: 레지스트리 탐색 + 보안 모듈 경로 확인

한컴 설치 여부 감지, 보안 DLL 경로 자동 탐색.

```rust
let hwp_path = HwpDetector::find_install_path()?;  // 레지스트리 fallback 탐색
let module_path = hwp_path.join("FilePathCheckerModule.dll");
```

**검증**: 설치된 한컴 버전과 DLL 경로가 정확히 탐지되는지

### Step 2: HWP 프로세스 띄우기 + 보안 모듈 등록

```rust
let hwp = HwpApp::new()?;
hwp.register_security_module(&module_path)?;
hwp.set_visible(true)?;
```

**검증**: HWP 창이 뜨고 보안 모듈 등록 성공

### Step 3: 문서 열기 + 텍스트 추출 (읽기)

```rust
hwp.open("C:\\temp\\test.hwp")?;
let text = hwp.get_text()?;  // GetTextFile("TEXT", "")
```

**검증**: HWP 파일의 텍스트가 정상 추출되는지

### Step 4: 텍스트 삽입 (쓰기)

HAction/HParameterSet 패턴으로 텍스트 삽입.

```rust
hwp.insert_text("안녕하세요")?;
hwp.save_as("C:\\temp\\output.hwp")?;
```

**검증**: 삽입된 텍스트가 저장 후 유지되는지

### Step 5: 실행 중 인스턴스 Attach

```rust
let hwp = HwpApp::get_active()?;  // GetActiveObject
let text = hwp.get_text()?;
```

**검증**: 수동으로 HWP를 열어두고 텍스트 읽기 성공

## 의존성

Phase 1과 동일 (`windows` crate) + 레지스트리 접근용 feature:

```toml
[dependencies]
windows = { version = "0.61", features = [
    "Win32_System_Com",
    "Win32_System_Ole",
    "Win32_System_Variant",
    "Win32_System_Registry",      # 레지스트리 탐색
    "Win32_Foundation",
] }
```

## 에러 처리 (Phase 1 추가분)

- 한컴 미설치 → `HwpNotInstalled` 에러
- 보안 모듈 등록 실패 → `SecurityModuleError` (경로/버전 정보 포함)
- 버전 감지 실패 → 지원 버전 목록과 함께 에러 반환

## 예제/테스트 파일

```
examples/
├── hwp_detect.rs          # Step 1: 한컴 설치 감지
├── hwp_hello.rs           # Step 2~3: 열기 + 텍스트 추출
├── hwp_write.rs           # Step 4: 텍스트 삽입 + 저장
└── hwp_attach.rs          # Step 5: 실행 중 인스턴스 연결
```
