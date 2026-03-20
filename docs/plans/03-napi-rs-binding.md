# Milestone 3: napi-rs 바인딩

## 목표

Rust COM 브릿지를 Node.js에서 사용할 수 있는 `.node` 네이티브 모듈로 빌드.
**개별 메서드 래핑 없이** generic dispatch 함수 3개만 노출.

## 핵심 작업

### 3-1. com-bridge-node crate 생성

```toml
[dependencies]
napi = { version = "2", features = ["napi8"] }
napi-derive = "2"
com-core = { path = "../com-core" }
```

### 3-2. 노출할 API (3 + 생명주기)

```rust
#[napi]
fn com_create(prog_id: String) -> Result<ExternalHandle>
// "Excel.Application" → handle 반환

#[napi]
fn com_get(handle: ExternalHandle, prop: String, args: Vec<JsVariant>) -> Result<JsVariant>
// dispatch.get() / get_by()

#[napi]
fn com_put(handle: ExternalHandle, prop: String, value: JsVariant) -> Result<()>
// dispatch.put()

#[napi]
fn com_call(handle: ExternalHandle, method: String, args: Vec<JsVariant>) -> Result<JsVariant>
// dispatch.call()

#[napi]
fn com_release(handle: ExternalHandle) -> Result<()>
// handle 해제
```

### 3-3. Handle 관리

- COM은 STA 스레드에서만 호출 가능
- napi-rs는 메인 스레드에서 실행 → STA 호환
- handle = HashMap<u32, DispatchObject> 또는 napi External

### 3-4. JsVariant 타입 매핑

```
JS string   ↔  Variant::String
JS number   ↔  Variant::F64 (정수면 I32)
JS boolean  ↔  Variant::Bool
JS null     ↔  Variant::Empty
JS object { _comHandle } ↔ Variant::Dispatch
```

## 검증

```js
const { comCreate, comGet, comPut, comCall } = require('./com-bridge-node.node');

const excel = comCreate("Excel.Application");
comPut(excel, "Visible", true);
const wbs = comGet(excel, "Workbooks");
const wb = comCall(wbs, "Add", []);
const sheet = comGet(wb, "ActiveSheet");
comPut(comCall(sheet, "Range", ["A1"]), "Value", "hello");
console.log(comGet(comCall(sheet, "Range", ["A1"]), "Value"));
// → "hello"
```
