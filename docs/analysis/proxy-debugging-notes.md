# JS Proxy + COM 브릿지 디버깅 노하우

실제 개발 중 발견한 이슈들과 해결법 정리.

---

## 1. `valueOf` / `Symbol.toPrimitive` 충돌

**증상**: `sheet.Range("E2").Value + " / " + sheet.Range("F2").Value`에서
`Number.prototype.valueOf requires that 'this' be a Number` 에러

**원인**: COM에서 숫자를 반환해도 Proxy로 감싸져 있어서, `+` 연산 시 JS가 `valueOf()`를 호출하는데 Proxy의 valueOf가 Number.prototype.valueOf와 충돌

**해결**: `createCallableProperty`의 get 트랩에 `Symbol.toPrimitive`, `valueOf`, `toString` 처리 추가

```js
if (subProp === Symbol.toPrimitive) {
  return () => {
    const result = bridge.comGet(parentHandle, prop);
    if (isComHandle(result)) return "[COM Object]";
    return result; // 원시값 그대로 반환
  };
}
if (subProp === "valueOf" || subProp === "toString") {
  return () => {
    const result = bridge.comGet(parentHandle, prop);
    if (isComHandle(result)) return "[COM Object]";
    return result;
  };
}
```

---

## 2. HWP HParameterSet 프로퍼티 직접 설정 실패

**증상**: `cs.Height = 2000`에서 `알 수 없는 이름입니다 (0x80020006)` (DISP_E_UNKNOWNNAME)

**원인**: Proxy의 callable property가 lazy하게 comGet을 호출하므로, 매번 다른 dispatch 참조를 가져올 수 있음. `GetDefault`가 호출된 객체와 property put 대상 객체가 다른 참조일 수 있음.

**해결**: HWP에서는 `CreateAction` + `SetItem` 방식 사용. 이 방식은 메서드 호출(call)이라 Proxy 체이닝 문제 없음.

```js
// ❌ Proxy 체이닝 문제 발생
var cs = hwp.HParameterSet.HCharShape;
hwp.HAction.GetDefault("CharShape", cs.HSet);
cs.Height = 2000; // 에러!

// ✅ CreateAction 방식 — 안정적
var csAct = hwp.CreateAction("CharShape");
var csSet = csAct.CreateSet();
csAct.GetDefault(csSet);
csSet.SetItem("Height", 2000); // SetItem은 메서드 호출이라 OK
csAct.Execute(csSet);
```

**교훈**: HWP COM은 property put보다 `SetItem()` 메서드 호출이 더 안정적. LLM에게 코드 생성 시킬 때 CreateAction 패턴을 기본으로 가이드해야 함.

---

## 3. `CreateAction` 반환값 null

**증상**: `hwp.CreateAction("FindReplace")`가 null 반환 → `CreateSet()` 호출 시 NPE

**원인**: HWP의 Action ID는 정확해야 함. `"FindReplace"`는 파라미터가 필요한 액션이 아니거나, 액션 ID가 다름.

**해결**: 찾기/바꾸기는 `"AllReplace"` 액션 사용

```js
// ❌ FindReplace는 CreateAction으로 안 될 수 있음
var findAct = hwp.CreateAction("FindReplace"); // null!

// ✅ AllReplace 사용
var frAct = hwp.CreateAction("AllReplace");
var frSet = frAct.CreateSet();
frAct.GetDefault(frSet);
frSet.SetItem("FindString", "찾을것");
frSet.SetItem("ReplaceString", "바꿀것");
frSet.SetItem("IgnoreMessage", 1);
frAct.Execute(frSet);
```

**참고**: 유효한 Action ID는 `docs/analysis/hwp-com-api-reference.md`의 Action ID 매핑 테이블 참조. 또는 한컴 디벨로퍼의 `ActionTable_2504.pdf` 참조.

---

## 4. Excel 워크북 없이 ActiveSheet 접근

**증상**: `excel.ActiveSheet.Range("A1")` → `Range is not a function`

**원인**: Excel 실행 직후에는 워크북이 없음. `ActiveSheet`가 null/empty를 반환하고, 그 위에 `.Range` 접근하면 실패.

**해결**: 반드시 `Workbooks.Add()` 또는 `Workbooks.Open()` 먼저 실행

```js
excel.Workbooks.Add(); // 이거 먼저!
excel.ActiveSheet.Range("A1").Value = "hello";
```

---

## 5. Proxy에서 `then` 프로퍼티 처리

**증상**: Proxy 객체를 `await`하거나 Promise 체인에 넣으면 무한 루프 또는 예기치 않은 동작

**원인**: JS의 Promise 시스템은 객체에 `.then`이 있으면 thenable로 인식하여 자동으로 처리하려 함

**해결**: Proxy의 get 트랩에서 `then` 접근 시 `undefined` 반환

```js
if (prop === "then") return undefined;
```

---

## 8. napi-rs v2 API 차이점

**증상**: `JsUnknown` 타입을 못 찾음, `OnceLock::get_or_try_init` unstable

**원인**: napi-rs v2의 `bindgen_prelude`에는 `JsUnknown`이 없음 (v1 스타일). `OnceLock::get_or_try_init`은 Rust stable에서 미지원.

**해결**:

```rust
// JsUnknown은 별도 import
use napi::JsUnknown;
use napi::NapiRaw;  // .raw() 메서드용

// OnceLock 대신 std::sync::Once
static COM_INIT: Once = Once::new();
fn ensure_init() {
    COM_INIT.call_once(|| {
        let rt = ComRuntime::init().expect("COM init failed");
        std::mem::forget(rt);  // 프로세스 종료까지 유지
    });
}
```

---

## 9. COM Variant → JS 변환 시 External 처리

**증상**: COM dispatch 객체를 JS로 반환할 때 `External::into_unknown()` 메서드 없음

**원인**: napi-rs v2에서 `External<T>`에는 `into_unknown()` 없음

**해결**: `env.create_external()`로 v1 스타일 JsExternal 생성 후 `into_unknown()`

```rust
Variant::Dispatch(disp) => {
    let external = env.create_external(disp, None)?;
    Ok(external.into_unknown())
}
```

---

## 6. PowerShell에서 빌드 스크립트 `cp` 호환성

**증상**: `package.json`의 `cp` 명령이 PowerShell에서 다르게 동작

**원인**: PowerShell의 `cp`는 `Copy-Item` alias. Unix `cp -r`이 아닌 `-Recurse` 플래그 사용.

**해결**: `shx cp`, Node.js `fs.cpSync`, 또는 bash 셸 명시로 회피.

---

## 7. Electron inline script 충돌 + CSP 경고

**증상**: `Identifier 'comBridge' has already been declared` + CSP 경고

**원인**: preload에서 `contextBridge.exposeInMainWorld('comBridge', ...)`로 노출한 뒤, HTML inline `<script>`에서 같은 이름으로 재선언.

**해결**: inline `<script>` 제거 → 별도 `renderer.js` 파일로 분리.

---

## Excel vs HWP 패턴 비교표

| 항목 | Excel | HWP |
|------|-------|-----|
| 프로퍼티 접근 | `sheet.Range("A1").Value = x` (직접 put) | `set.SetItem("Height", x)` (메서드 호출) |
| 실행 | 즉시 반영 | `action.Execute(set)` 호출 필요 |
| 기본 흐름 | get → put 체이닝 | CreateAction → CreateSet → GetDefault → SetItem → Execute |
| Proxy 호환성 | 높음 (get/set 직접 대응) | 낮음 (SetItem 메서드 사용 권장) |
| 문서 생성 | `excel.Workbooks.Add()` | `hwp.Run("FileNew")` |
| 찾기바꾸기 | `Range.Replace()` | `CreateAction("AllReplace")` |

---

## LLM 코드 생성 가이드라인 (이 노하우 기반)

1. **Excel**: `Workbooks.Add()` 먼저 → 자연스러운 COM 문법 OK
2. **HWP**: `CreateAction` + `SetItem` 패턴 기본 → `HParameterSet` 직접 property put 금지
3. **찾기바꾸기**: HWP는 `AllReplace` 액션만 사용 (`FindReplace`는 CreateAction 불가)
4. **값 읽기**: 문자열 연결 시 Proxy 자동 변환됨 (Symbol.toPrimitive)
5. **에러 시**: 세이브포인트에서 자동 복원되므로 문서 상태 걱정 불필요
6. **Action ID 확인**: null 반환되면 ID가 틀린 것. ActionTable PDF 또는 CLI로 검증
