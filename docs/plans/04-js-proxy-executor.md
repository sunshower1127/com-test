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

### 4-3. 세이브포인트 (자동 롤백)

COM은 트랜잭션이 없으므로 LLM 코드가 실패하면 문서가 중간 상태로 남는다.
실행기가 **자동으로** 세이브포인트를 관리하여 LLM은 비즈니스 로직에만 집중.

```js
async function executeWithSavepoint(llmCode, sandbox, appType) {
  const tempPath = `${os.tmpdir()}/savepoint_${Date.now()}.tmp`;

  // 1. 세이브포인트 생성 (LLM 모르게 자동)
  if (appType === 'excel') {
    comCall(sandbox.excel._handle, "SaveCopyAs", [tempPath]);
  } else if (appType === 'hwp') {
    comCall(sandbox.hwp._handle, "SaveAs", [tempPath, "HWP", ""]);
  }

  // 2. LLM 코드 실행
  const result = runInSandbox(llmCode, sandbox);

  // 3. 실패 시 복원
  if (!result.success) {
    if (appType === 'excel') {
      comPut(sandbox.excel._handle, "DisplayAlerts", false);
      comCall(sandbox.excel._handle, "Close", [false]);
      // workbooks.open으로 복원
    } else if (appType === 'hwp') {
      comCall(sandbox.hwp._handle, "Clear", [1]);
      comCall(sandbox.hwp._handle, "Open", [tempPath]);
    }
  }

  fs.unlinkSync(tempPath);
  return result;
}
```

**설계 원칙:**
- LLM은 세이브포인트 존재를 모름 — 순수 비즈니스 코드만 생성
- 인프라 레벨에서 자동 처리 → LLM이 빼먹을 가능성 0
- 복원 코드는 우리가 작성 → 복원 자체의 에러 가능성 최소화

### 4-4. 에러 포맷팅

LLM 재시도를 위한 구조화된 에러:

```js
{
  success: false,
  error: "Cannot read property 'Range' of undefined",
  line: 3,
  code: "const val = excel.ActiveShet.Range('A1').Value;",
  hint: "property 'ActiveShet' returned undefined — 오타?",
  logs: ["step1 done", "step2 done"],  // 에러 전까지의 로그
  restored: true,  // 세이브포인트에서 복원됨
}
```

## 검증

```js
// 1. 정상 실행
const result = executeCode(`
  const sheet = excel.ActiveSheet;
  sheet.Range("A1").Value = "hello";
  result = sheet.Range("A1").Value;
`, { excel: excelHandle });
assert(result.success === true);
assert(result.result === "hello");

// 2. 에러 → 자동 복원 확인
const result2 = executeWithSavepoint(`
  excel.ActiveSheet.Range("A1").Value = "modified";
  throw new Error("중간에 에러");
`, sandbox, 'excel');
assert(result2.success === false);
assert(result2.restored === true);
// A1은 "modified"가 아닌 원래 값으로 복원되어야 함
```
