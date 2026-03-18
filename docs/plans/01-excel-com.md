# Phase 1: Excel COM 자동화 (Rust)

## 목적

Rust `windows` crate로 Excel COM Automation 패턴을 확립한다.
여기서 잡은 구조(STA 초기화, IDispatch 호출, 에러 처리)를 Phase 2 HWP에 그대로 적용한다.

## 사전 조건

- Windows 11
- MS Office (Excel) 설치됨
- Rust toolchain (`stable-x86_64-pc-windows-msvc`)

## 핵심 개념

### COM 호출 흐름

```
CoInitializeEx(COINIT_APARTMENTTHREADED)   // 1. STA 초기화
    │
CoCreateInstance("Excel.Application")       // 2. COM 객체 생성
    │
IDispatch::Invoke(...)                      // 3. 메서드/프로퍼티 호출
    │
Release + CoUninitialize                    // 4. 정리
```

### Excel COM 객체 계층

```
Excel.Application
├── .Visible          // bool - UI 표시 여부
├── .Workbooks
│   ├── .Add()        // 새 워크북
│   ├── .Open(path)   // 기존 파일 열기
│   └── [index]
│       └── Workbook
│           ├── .Sheets[index] → Worksheet
│           │   ├── .Cells(row, col) → Range
│           │   │   ├── .Value     // 읽기/쓰기
│           │   │   └── .Formula   // 수식
│           │   ├── .Range("A1:B5") → Range
│           │   └── .Name
│           ├── .SaveAs(path)
│           ├── .Save()
│           └── .Close()
└── .Quit()
```

## 구현 단계

### Step 1: 프로젝트 세팅 + COM 초기화

`com-core` crate 생성. STA 초기화/해제를 안전하게 감싸는 구조 만들기.

```rust
// 목표 API
let _com = ComRuntime::init()?;  // STA 초기화, Drop 시 CoUninitialize
```

**검증**: `CoInitializeEx` 호출 성공 확인

### Step 2: Excel 프로세스 띄우기

`excel-bridge` crate 생성. `Excel.Application` COM 객체 생성 및 Visible 설정.

```rust
let excel = ExcelApp::new()?;
excel.set_visible(true)?;   // 화면에 Excel이 보여야 함
// Drop 시 excel.quit() + Release
```

**검증**: Excel 창이 뜨고, Drop 시 정상 종료되는지 확인

### Step 3: 워크북/시트 조작

워크북 생성, 열기, 시트 접근.

```rust
let wb = excel.workbooks().add()?;
let sheet = wb.sheet(1)?;
```

**검증**: 빈 워크북이 생성되는지 확인

### Step 4: 셀 읽기/쓰기

셀 값 읽기, 쓰기, 수식 설정.

```rust
sheet.cell(1, 1).set_value("Hello")?;
sheet.cell(1, 2).set_formula("=A1 & \" World\"")?;
let val = sheet.cell(1, 2).value()?;  // "Hello World"
```

**검증**: 값이 정상적으로 읽히고 수식이 계산되는지 확인

### Step 5: 파일 저장/열기

SaveAs, Open, Close.

```rust
wb.save_as("C:\\temp\\test.xlsx")?;
wb.close()?;

let wb2 = excel.workbooks().open("C:\\temp\\test.xlsx")?;
let val = wb2.sheet(1)?.cell(1, 1).value()?;  // "Hello"
```

**검증**: 저장 후 다시 열어서 값이 유지되는지 확인

### Step 6: 이미 열린 Excel에 Attach

`GetActiveObject`로 기존 인스턴스에 연결.

```rust
let excel = ExcelApp::get_active()?;  // 이미 실행 중인 Excel에 연결
let wb = excel.active_workbook()?;
```

**검증**: 수동으로 Excel을 열어두고, attach해서 내용 읽기 성공

## 의존성

```toml
[dependencies]
windows = { version = "0.61", features = [
    "Win32_System_Com",
    "Win32_System_Ole",
    "Win32_System_Variant",
    "Win32_Foundation",
] }
```

## 에러 처리 전략

- COM `HRESULT` → Rust `Result<T, ComError>` 변환
- 다이얼로그 블로킹 방지: `Application.DisplayAlerts = false`
- 좀비 프로세스 방지: `Drop` trait으로 `Quit()` + `Release` 보장

## 예제/테스트 파일

```
examples/
├── excel_hello.rs         # Step 2~4 통합: 셀 쓰기 → 읽기
├── excel_file_io.rs       # Step 5: 저장/열기
└── excel_attach.rs        # Step 6: 실행 중 인스턴스 연결
```
