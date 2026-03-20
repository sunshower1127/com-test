# COM Bridge 프로젝트 개요

## 목표

Windows COM Automation을 통해 한컴오피스(HWP)와 MS Office(Excel/Word)를 외부에서 제어하는 브릿지 모듈을 만든다.
최종적으로 Electron 앱(Wissly)에서 `napi-rs` + `utilityProcess` 구조로 통합한다.

## 기술 스택

| 레이어 | 기술 | 비고 |
|--------|------|------|
| COM 바인딩 | Rust + `windows` crate | Microsoft 공식 crate |
| Node.js 바인딩 | napi-rs | `.node` 모듈로 빌드 |
| 프로세스 격리 | Electron `utilityProcess` | 크래시 격리, 서명 추가 없음 |
| IPC | `MessagePort` (postMessage) | Electron 내장 |

## 아키텍처

```
Renderer ──IPC──► Main Process ──postMessage──► utilityProcess (bridge.js)
                                                      │
                                                 .node (napi-rs)
                                                      │
                                                 Rust COM 호출
                                                      │
                                          ┌───────────┴───────────┐
                                          │                       │
                                    HWP COM Server          Office COM Server
                                   (HWPFrame.HwpObject)    (Excel.Application)
```

## 프로젝트 구조 (목표)

```
com-test/
├── crates/
│   ├── com-core/            # COM 초기화, 공통 유틸 (STA, 에러 처리)
│   ├── excel-bridge/        # Excel COM 자동화
│   ├── hwp-bridge/          # HWP COM 자동화
│   └── com-bridge-node/     # napi-rs 바인딩 (Phase 3)
├── electron/                # Electron 앱 (Phase 4)
│   ├── src/
│   │   ├── main/            # Main process
│   │   ├── renderer/        # Renderer
│   │   └── bridge/          # utilityProcess 엔트리
│   └── package.json
├── docs/
│   ├── plans/               # 계획 문서
│   └── references/          # 참고 자료
├── Cargo.toml               # Workspace root
└── examples/                # 독립 실행 가능한 예제/테스트
```

## 로드맵

| Phase | 내용 | 상세 문서 |
|-------|------|-----------|
| 1 | Excel COM 자동화 (Rust) | [01-excel-com.md](./01-excel-com.md) |
| 2 | HWP COM 자동화 (Rust) | [02-hwp-com.md](./02-hwp-com.md) |
| 3 | napi-rs 바인딩 | [03-napi-rs.md](./03-napi-rs.md) |
| 4 | Electron 연동 | [04-electron.md](./04-electron.md) |

## 설계 원칙

- **크래시 격리 우선**: COM 호출은 반드시 메인 프로세스와 분리
- **STA 준수**: COM 초기화 스레드에서만 호출, Rust에서 명시적으로 제어
- **단일 바이너리**: 런타임 의존성 없이 `.node` 또는 `.exe` 하나로 배포
- **점진적 통합**: 각 Phase가 독립적으로 동작 확인 가능해야 함
