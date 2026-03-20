# Milestone 4: JS Proxy + VM 실행기

## 목표

napi-rs의 generic dispatch를 JS Proxy로 감싸서 자연스러운 COM 문법을 제공하고,
LLM이 생성한 코드를 vm 샌드박스에서 안전하게 실행.

## 핵심 작업

### 4-1. createComProxy 구현

```js
function createComProxy(handle) {
  return new Proxy(function() {}, {
    get(_, prop) {
      if (prop === '_handle') return handle;
      const result = comGet(handle, prop);
      if (result && result._comHandle) {
        return createComProxy(result._comHandle);
      }
      return result;
    },
    set(_, prop, value) {
      comPut(handle, prop, value);
      return true;
    },
    apply(_, thisArg, args) {
      // Range("A1") 같은 호출 처리
      // → 부모 객체의 메서드 호출로 변환
    }
  });
}
```

과제: `sheet.Range("A1")` — Range가 get인 동시에 call.
→ get 시점에 함수를 반환하되, 인자 없이 접근하면 프로퍼티로 동작하는 방식 필요.

### 4-2. VM 샌드박스 실행기

```js
const vm = require('vm');

function executeCode(code, apps) {
  const sandbox = {
    excel: apps.excel ? createComProxy(apps.excel) : undefined,
    hwp: apps.hwp ? createComProxy(apps.hwp) : undefined,
    console: { log: (...args) => logs.push(args) },
    result: null,  // LLM이 결과를 담을 변수
  };

  const logs = [];

  try {
    vm.runInNewContext(code, sandbox, {
      timeout: 30000,        // 30초 타임아웃
      displayErrors: true,
    });
    return {
      success: true,
      result: sandbox.result,
      logs,
    };
  } catch (e) {
    return {
      success: false,
      error: e.message,
      stack: e.stack,
      logs,
      code,
    };
  }
}
```

### 4-3. 에러 포맷팅

LLM 재시도를 위한 구조화된 에러:

```js
{
  success: false,
  error: "Cannot read property 'Range' of undefined",
  line: 3,
  code: "const val = excel.ActiveShet.Range('A1').Value;",
  hint: "property 'ActiveShet' returned undefined — 오타?",
  logs: ["step1 done", "step2 done"],  // 에러 전까지의 로그
}
```

## 검증

```js
const result = executeCode(`
  const sheet = excel.ActiveSheet;
  sheet.Range("A1").Value = "hello";
  result = sheet.Range("A1").Value;
`, { excel: excelHandle });

assert(result.success === true);
assert(result.result === "hello");
```
