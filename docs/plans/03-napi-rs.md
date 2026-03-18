# Phase 3: napi-rs 바인딩

## 목적

Phase 1~2에서 만든 Rust COM 브릿지를 napi-rs로 감싸서
Node.js/Electron에서 사용 가능한 `.node` 모듈로 빌드한다.

## 사전 조건

- Phase 1, 2 완성
- Node.js 설치
- napi-rs CLI (`@napi-rs/cli`)

## 구조

```
com-bridge-node/          # napi-rs crate
├── src/
│   └── lib.rs            # #[napi] 함수들 — JS에 노출되는 인터페이스
├── Cargo.toml
├── package.json           # npm 패키지 정의
└── build.rs
```

빌드 결과: `com-bridge.{platform}.node`

## JS 인터페이스 설계

```typescript
// TypeScript 타입 정의 (자동 생성됨)

// Excel
export function excelOpen(path: string): ExcelHandle
export function excelCreate(): ExcelHandle
export function excelGetCellValue(handle: ExcelHandle, sheet: number, row: number, col: number): string | number | null
export function excelSetCellValue(handle: ExcelHandle, sheet: number, row: number, col: number, value: string | number): void
export function excelSave(handle: ExcelHandle, path?: string): void
export function excelClose(handle: ExcelHandle): void

// HWP
export function hwpOpen(path: string): HwpHandle
export function hwpGetText(handle: HwpHandle): string
export function hwpInsertText(handle: HwpHandle, text: string): void
export function hwpSave(handle: HwpHandle, path?: string): void
export function hwpClose(handle: HwpHandle): void

// 공통
export function isOfficeInstalled(): boolean
export function isHwpInstalled(): boolean
```

## 구현 단계

### Step 1: napi-rs 프로젝트 초기화

```bash
npx @napi-rs/cli new com-bridge-node
```

### Step 2: Excel 함수 바인딩

Phase 1의 `excel-bridge`를 `#[napi]` 함수로 래핑.

```rust
#[napi]
pub fn excel_create() -> Result<ExcelHandle> {
    let app = ExcelApp::new()?;
    Ok(ExcelHandle::new(app))
}
```

### Step 3: HWP 함수 바인딩

Phase 2의 `hwp-bridge`를 동일 패턴으로 래핑.

### Step 4: Handle 관리

COM 객체는 STA 스레드에 묶여야 하므로,
napi-rs 측에서 handle → 내부 HashMap으로 관리하고,
실제 COM 호출은 전용 STA 스레드로 dispatch.

```
JS 호출 → #[napi] 함수 → channel.send(command)
                              │
                         STA Thread (COM 전용)
                              │
                         COM 호출 → 결과 반환
```

### Step 5: 빌드 + 테스트

```bash
npm run build
node -e "const b = require('./com-bridge.node'); console.log(b.isOfficeInstalled())"
```

## 주의사항

- **STA 스레드 관리**: napi-rs의 `#[napi]` 함수는 libuv 스레드에서 호출됨.
  COM 호출을 직접 하면 안 되고, 별도 STA 스레드로 위임해야 함.
- **Handle 수명**: JS GC와 COM 객체 수명이 다름.
  명시적 `close()` 호출 or PoC 단계에선 수동 관리.
- **에러 변환**: `ComError` → napi-rs `Error`로 깔끔하게 매핑.
