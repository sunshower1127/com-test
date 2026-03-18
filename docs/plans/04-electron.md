# Phase 4: Electron 연동

## 목적

Phase 3에서 만든 `.node` 모듈을 Electron `utilityProcess`에서 로드하여
크래시 격리된 COM 브릿지를 완성한다.

## 사전 조건

- Phase 3 `.node` 빌드 완성
- Electron 22+ (utilityProcess 지원)

## 아키텍처

```
┌─────────────────────────────────────────────────┐
│ Electron App                                    │
│                                                 │
│  Renderer ──preload/IPC──► Main Process         │
│                              │                  │
│                         utilityProcess.fork()   │
│                              │                  │
│                    ┌─────────▼──────────┐       │
│                    │  bridge.js         │       │
│                    │  (별도 프로세스)      │       │
│                    │                    │       │
│                    │  .node (napi-rs)   │       │
│                    │       │            │       │
│                    │   COM 호출          │       │
│                    └────────────────────┘       │
└─────────────────────────────────────────────────┘
```

## 구현 단계

### Step 1: Electron 프로젝트 세팅

```
electron/
├── src/
│   ├── main/
│   │   └── index.ts          # Main process
│   ├── renderer/
│   │   └── index.html        # 간단한 UI
│   └── bridge/
│       └── bridge.js         # utilityProcess 엔트리
├── package.json
└── electron-builder.yml
```

### Step 2: bridge.js (utilityProcess 엔트리)

```js
// bridge.js — utilityProcess에서 실행됨
const { HwpBridge, ExcelBridge } = require('./com-bridge.node')

process.parentPort.on('message', async (e) => {
    const { id, target, action, args } = e.data
    try {
        let result
        if (target === 'excel') {
            result = ExcelBridge[action](...args)
        } else if (target === 'hwp') {
            result = HwpBridge[action](...args)
        }
        process.parentPort.postMessage({ id, result })
    } catch (err) {
        process.parentPort.postMessage({ id, error: err.message })
    }
})
```

### Step 3: Main process에서 utilityProcess 관리

```ts
// main/index.ts
const bridge = utilityProcess.fork(path.join(__dirname, '../bridge/bridge.js'), [], {
    serviceName: 'com-bridge'
})

// 요청-응답 패턴 (id 기반 매칭)
let reqId = 0
const pending = new Map()

function callBridge(target, action, ...args) {
    return new Promise((resolve, reject) => {
        const id = reqId++
        pending.set(id, { resolve, reject })
        bridge.postMessage({ id, target, action, args })
    })
}

bridge.on('message', ({ id, result, error }) => {
    const p = pending.get(id)
    if (p) {
        pending.delete(id)
        error ? p.reject(new Error(error)) : p.resolve(result)
    }
})

// 크래시 시 자동 재시작
bridge.on('exit', (code) => {
    if (code !== 0) restartBridge()
})
```

### Step 4: Renderer ↔ Main IPC

```ts
// preload.ts
const { ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('comBridge', {
    excel: {
        open: (path) => ipcRenderer.invoke('bridge:excel:open', path),
        getCellValue: (handle, sheet, row, col) =>
            ipcRenderer.invoke('bridge:excel:getCellValue', handle, sheet, row, col),
        // ...
    },
    hwp: {
        open: (path) => ipcRenderer.invoke('bridge:hwp:open', path),
        getText: (handle) => ipcRenderer.invoke('bridge:hwp:getText', handle),
        // ...
    }
})
```

### Step 5: 간단한 테스트 UI

- 파일 선택 → Excel/HWP 열기
- 내용 읽기/쓰기
- 상태 표시 (연결됨/끊김/재시작 중)

## 서명 관련

| 파일 | 서명 필요 |
|------|----------|
| electron.exe | O (기존) |
| .node (napi-rs) | X (Electron 내부 로드) |
| bridge.js | X (스크립트) |

별도 `.exe` 가 없으므로 기존 Electron 서명만으로 충분.

## 에러 처리

- utilityProcess 크래시 → 자동 재시작 + UI에 알림
- COM 타임아웃 → 요청별 타임아웃 설정 (다이얼로그 블로킹 대비)
- Office/HWP 미설치 → 기능 비활성화 + 사용자 안내
