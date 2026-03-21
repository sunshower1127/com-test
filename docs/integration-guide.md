# COM Bridge 통합 가이드

다른 Electron 프로젝트에서 COM Bridge를 사용하여 Office 자동화를 구현하는 방법.

---

## 개요

이 프로젝트는 LLM이 생성한 JS 코드를 Electron 앱에서 실행하여 Excel, Word, PPT, 한글(HWP)을 자동 조작하는 시스템입니다.

```
LLM ──js-com 코드──▶ Electron App ──napi-rs──▶ COM ──▶ Office 앱
                       ◀──결과/에러──
```

---

## 필요한 파일

### 1. 네이티브 모듈

```
native/com_bridge_node.node
```

napi-rs로 빌드된 네이티브 모듈. COM 통신을 담당합니다.

**빌드 방법:**
```bash
cd crates/com-bridge-node
npm run build    # 또는 cargo build
```

빌드 결과물: `target/debug/com_bridge_node.node` (또는 release)

이 `.node` 파일을 당신의 Electron 프로젝트에 복사하세요.

### 2. Worker 코드

```
electron/dist/worker/
├── index.js       # Worker 엔트리포인트
├── proxy.js       # JS Proxy (COM 객체를 JS처럼 사용)
├── executor.js    # VM 샌드박스 실행기
```

이 3개 파일을 당신의 프로젝트에 복사하세요.

### 3. 프롬프트 파일

```
docs/prompts/
├── office-automation-skill.md    # LLM 시스템 프롬프트 (스킬 활성화)
├── hwp-cheatsheet-compact.md     # HWP 전용 cheat sheet
```

---

## 통합 방법

### Step 1: Worker 프로세스 설정

Electron의 `utilityProcess` 또는 `child_process.fork`로 Worker를 실행합니다.

```js
// main process에서
const { utilityProcess } = require('electron');

const worker = utilityProcess.fork(
  path.join(__dirname, 'worker/index.js')
);

// Worker에 메시지 전송
worker.postMessage({ type: 'launch', id: '1', app: 'excel' });

// Worker 응답 수신
worker.on('message', (msg) => {
  console.log(msg); // { type: 'status', excel: true, hwp: false, ... }
});
```

### Step 2: Worker IPC 프로토콜

**요청 타입:**

```typescript
// 앱 실행
{ type: 'launch', id: string, app: 'excel' | 'word' | 'ppt' | 'hwp' }

// 앱 종료
{ type: 'quit', id: string, app: 'excel' | 'word' | 'ppt' | 'hwp' }

// 코드 실행
{ type: 'execute', id: string, code: string }
```

**응답 타입:**

```typescript
// 앱 상태
{ type: 'status', excel: boolean, hwp: boolean, word: boolean, ppt: boolean }

// 실행 성공
{ type: 'result', id: string, success: true, result: any, logs: string[] }

// 실행 에러
{
  type: 'error', id: string, success: false,
  error: string,      // 에러 메시지
  stack: string,      // 스택트레이스
  line: number,       // 에러 줄 번호
  code: string,       // 실행한 원본 코드
  logs: string[],     // console.log 출력
  restored: boolean   // (사용 안 함, 항상 false)
}
```

### Step 2.5: 자동 Launch (권장)

`execute` 요청 전에 코드에서 사용하는 앱을 감지하여 자동 launch하면 사용자/LLM이 launch를 신경 쓸 필요 없습니다.

```js
// Worker 내부 또는 메인 프로세스에서 execute 전에 체크
function autoLaunch(code, apps) {
  if (code.includes('excel') && !apps.excel) {
    handleMessage({ type: 'launch', id: '0', app: 'excel' });
  }
  if (code.includes('hwp') && !apps.hwp) {
    handleMessage({ type: 'launch', id: '0', app: 'hwp' });
  }
  if (code.includes('word') && !apps.word) {
    handleMessage({ type: 'launch', id: '0', app: 'word' });
  }
  if (code.includes('ppt') && !apps.ppt) {
    handleMessage({ type: 'launch', id: '0', app: 'ppt' });
  }
}
```

이렇게 하면 LLM이 생성한 코드가 `excel.Workbooks.Add()`로 시작하면 Excel이 자동으로 실행됩니다.

### Step 3: LLM 연동 플로우

```
┌──────────────────────────────────────────────────────────────┐
│  1. 사용자가 "Office 스킬 활성화" 버튼 클릭                    │
│     → office-automation-skill.md 내용을 LLM에게 전송           │
│                                                                │
│  2. (선택) HWP 작업이면 "HWP Cheat Sheet" 버튼 클릭            │
│     → hwp-cheatsheet-compact.md 내용을 LLM에게 추가 전송       │
│                                                                │
│  3. 사용자: "A1~A10에 1~10 채워줘"                             │
│     → LLM이 ```js-com 코드 생성                               │
│                                                                │
│  4. 앱이 ```js-com 코드 블록 감지                              │
│     → Worker에 { type: 'execute', code: '...' } 전송           │
│                                                                │
│  5. Worker 실행 결과를 LLM에게 자동 전송                        │
│     → 성공: "실행 결과: ..."                                    │
│     → 에러: "에러 발생: ... Line: 15 ..."                      │
│                                                                │
│  6. LLM이 판단                                                  │
│     → 성공: "완료했습니다" 또는 다음 단계 코드                  │
│     → 에러: 수정된 코드 생성 (에러 지점부터 이어서)             │
└──────────────────────────────────────────────────────────────┘
```

### Step 4: js-com 코드 블록 감지

LLM 응답에서 ` ```js-com ` 코드 블록을 파싱합니다:

```js
function extractJsComCode(llmResponse) {
  var match = llmResponse.match(/```js-com\n([\s\S]*?)```/);
  return match ? match[1].trim() : null;
}
```

### Step 5: 실행 결과를 LLM에게 피드백

```js
function formatResultForLLM(response) {
  if (response.success) {
    var parts = [];
    if (response.logs.length) {
      parts.push('--- logs ---\n' + response.logs.join('\n'));
    }
    if (response.result !== null && response.result !== undefined) {
      parts.push('--- result ---\n' + JSON.stringify(response.result, null, 2));
    } else {
      parts.push('--- OK (no return value) ---');
    }
    return parts.join('\n\n');
  } else {
    var parts = ['Error: ' + response.error];
    if (response.line) parts.push('Line: ' + response.line);
    if (response.logs.length) parts.push('\n--- logs ---\n' + response.logs.join('\n'));
    if (response.stack) parts.push('\n--- stack ---\n' + response.stack);
    return parts.join('\n');
  }
}
```

---

## 앱 실행 순서

Worker가 실행되면 자동으로 COM이 초기화됩니다.
Office 앱은 사용자가 필요할 때 launch합니다.

```
Worker 시작 → COM 초기화 (자동)
                ↓
사용자: "엑셀에서 뭐 해줘"
                ↓
launch excel → Excel 프로세스 실행 + Visible=true
                ↓
execute 코드 → VM 샌드박스에서 실행 → COM 호출 → Excel 조작
                ↓
quit excel → DisplayAlerts=false → Quit → 프로세스 종료
```

---

## 앱별 특이사항

### Excel / Word
- Launch 시 바로 Visible=true
- `Workbooks.Add()` / `Documents.Add()` 로 새 문서 생성
- Quit 시 `DisplayAlerts=false` → `Quit()`

### PowerPoint
- Launch 시 **Visible 설정 안 함** (프레젠테이션 없으면 에러)
- `Presentations.Add()` 하면 자동으로 창 보임
- 포커스 필요하면 `ppt.Activate()`

### HWP (한글)
- Launch 시 Visible 토글 (False→True, 2024 호환 워크어라운드)
- **CreateAction 패턴 필수** — 직접 property put 사용 금지
- Quit 시 `Clear(1)` → `Quit()`
- 보안 모듈 미설정 시 파일 열기/저장에서 팝업 발생 가능

---

## 지원 환경

- **OS**: Windows만 (COM은 Windows 전용)
- **HWP**: 한글 2018 이상 (HwpAutomation, `HWPFrame.HwpObject`)
- **MS Office**: 설치된 버전 (COM 등록되어 있으면 동작)
- **Electron**: 28+ 권장
- **Node.js**: 18+ (vm 모듈 사용)

---

## 프롬프트 전송

### 원격 URL (GitHub Raw)

프롬프트를 로컬에 복사하지 않고 서버에서 직접 가져옵니다. 프롬프트 수정 시 git push만 하면 클라이언트가 자동으로 최신 버전을 받습니다.

```
Office 스킬:
https://raw.githubusercontent.com/sunshower1127/com-test/refs/heads/main/docs/prompts/office-automation-skill.md

HWP Cheat Sheet:
https://raw.githubusercontent.com/sunshower1127/com-test/refs/heads/main/docs/prompts/hwp-cheatsheet-compact.md
```

### 사용 예시

```js
// 프롬프트 fetch → LLM에 전송
async function sendSkillPrompt(type) {
  const urls = {
    office: 'https://raw.githubusercontent.com/sunshower1127/com-test/refs/heads/main/docs/prompts/office-automation-skill.md',
    hwp: 'https://raw.githubusercontent.com/sunshower1127/com-test/refs/heads/main/docs/prompts/hwp-cheatsheet-compact.md',
  };

  const res = await fetch(urls[type]);
  const prompt = await res.text();

  // LLM에게 user message로 전송
  sendToLLM({ role: 'user', content: prompt });
}
```

### 버튼 연결

| 버튼 | 동작 | 언제 |
|------|------|------|
| **Office 스킬 활성화** | `sendSkillPrompt('office')` | 오피스 자동화 시작 시 1회 |
| **HWP Cheat Sheet** | `sendSkillPrompt('hwp')` | 한글 작업 시 추가 전송 |

마지막에 "이 프롬프트에 대한 대답은 할 필요 없어"가 포함되어 있어 LLM이 불필요한 응답을 하지 않습니다.

---

## 파일 구조 요약

```
당신의 Electron 프로젝트/
├── native/
│   └── com_bridge_node.node    # napi-rs 네이티브 모듈 (복사)
├── worker/
│   ├── index.js                # Worker 엔트리 (복사)
│   ├── proxy.js                # JS Proxy (복사)
│   └── executor.js             # VM 샌드박스 (복사)
└── (당신의 main/renderer 코드)
    # 프롬프트는 GitHub Raw에서 fetch — 로컬 복사 불필요
```
