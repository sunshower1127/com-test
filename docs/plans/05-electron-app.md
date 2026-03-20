# Milestone 5: Electron 앱

## 목표

COM 브릿지 + VM 실행기를 Electron 앱으로 통합.
utilityProcess에서 COM을 실행하고, 메인/렌더러와 IPC로 통신.

## 핵심 작업

### 5-1. 프로젝트 구조

```
electron/
├── package.json
├── src/
│   ├── main/
│   │   ├── index.ts            # 메인 프로세스
│   │   └── com-worker.ts       # utilityProcess fork 관리
│   ├── worker/
│   │   ├── index.ts            # utilityProcess 엔트리
│   │   ├── proxy.ts            # createComProxy
│   │   └── executor.ts         # VM 실행기
│   ├── renderer/
│   │   └── ...                 # UI
│   └── shared/
│       └── types.ts            # IPC 메시지 타입
└── native/
    └── com-bridge-node.node    # Milestone 3 산출물
```

### 5-2. 프로세스 구조

```
메인 프로세스 (main)
  ├── 렌더러 (renderer) — UI, 채팅 인터페이스
  └── utilityProcess (worker) — COM 브릿지 + VM 실행
       └── .node 네이티브 모듈 로드
```

### 5-3. IPC 메시지 흐름

```
Renderer → Main → Worker:  { type: "execute", code: "..." }
Worker → Main → Renderer:  { type: "result", success: true, result: ... }
Worker → Main → Renderer:  { type: "error", error: "...", stack: "..." }
Worker → Main → Renderer:  { type: "log", args: [...] }
```

### 5-4. 앱 생명주기

1. Electron 시작 → utilityProcess fork
2. Worker: COM 초기화 (STA)
3. Worker: .node 모듈 로드
4. User → Renderer: "A1에 hello 넣어줘"
5. Renderer → Main → LLM API: 코드 생성 요청
6. LLM 응답 → Main → Worker: 코드 실행
7. Worker → Main → Renderer: 결과/에러

### 5-5. UI (최소)

- 채팅 인터페이스 (입력 + 응답)
- 실행 로그 패널
- 앱 상태 표시 (Excel/HWP 실행 중 여부)

## 검증

1. Electron 앱 실행 → utilityProcess에서 COM 정상 동작
2. Renderer에서 코드 전송 → Worker 실행 → 결과 렌더링
3. Excel/HWP 창이 뜨고 조작되는 것 확인
