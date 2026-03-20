# Office COM Bridge — 프로젝트 로드맵

## 목표

LLM이 JS 코드를 생성하면, Electron 앱에서 안전하게 실행하여
Excel / HWP / Word / PPT 등 오피스 문서를 자동 조작하는 시스템.

## 아키텍처

```
┌──────────┐   JS 코드    ┌─────────────────────────────────┐
│   LLM    │ ──────────▶ │  Electron App                    │
│          │             │                                   │
│          │ ◀────────── │  ┌───────────────────────────┐   │
│  (retry) │  에러/결과   │  │ utilityProcess             │   │
└──────────┘             │  │                             │   │
                         │  │  vm.runInNewContext(code, { │   │
                         │  │    excel: Proxy → napi-rs   │   │
                         │  │    hwp:   Proxy → napi-rs   │   │
                         │  │  })                         │   │
                         │  │         │                   │   │
                         │  │         ▼                   │   │
                         │  │  napi-rs (.node)            │   │
                         │  │  comGet / comPut / comCall   │   │
                         │  │         │                   │   │
                         │  │         ▼                   │   │
                         │  │  Rust COM Bridge            │   │
                         │  │  (com-core + bridges)       │   │
                         │  └───────────────────────────┘   │
                         └─────────────────────────────────┘
                                      │ COM IDispatch
                                      ▼
                              Excel / HWP / Word / PPT
```

## 마일스톤

| # | 마일스톤 | 핵심 산출물 | 상태 |
|---|---------|------------|------|
| 1 | [Rust COM 기반](01-rust-com-foundation.md) | com-core, excel-bridge, hwp-bridge, com-cli | ✅ 완료 |
| 2 | [Type 탐색 + HWP Cheat Sheet](02-type-introspection.md) | ITypeInfo 메서드 열거, HWP cheat sheet (Office는 스킵) | 🔲 |
| 3 | [napi-rs 바인딩](03-napi-rs-binding.md) | generic dispatch .node 모듈 | 🔲 |
| 4 | [JS Proxy + VM 실행기](04-js-proxy-executor.md) | createComProxy, vm 샌드박스, 에러 핸들링 | 🔲 |
| 5 | [Electron 앱](05-electron-app.md) | utilityProcess, IPC, UI | 🔲 |
| 6 | [LLM 연동](06-llm-integration.md) | 코드 생성, 에러 retry 루프, cheat sheet 주입 | 🔲 |

## 핵심 설계 원칙

1. **래핑 최소화** — 개별 메서드를 감싸지 않음. comGet/comPut/comCall 3개로 모든 COM 호출 처리
2. **JS Proxy** — COM의 late-binding과 JS Proxy의 동적 접근이 1:1 대응
3. **vm 샌드박스** — LLM 코드를 격리 실행. sandbox에 넣은 것만 접근 가능
4. **에러 피드백 루프** — 실패 시 에러+코드를 LLM에 재전송하여 자동 수정
5. **Cheat Sheet** — 전체 문서 대신 앱별 요약으로 LLM 가이드
