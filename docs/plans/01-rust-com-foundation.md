# Milestone 1: Rust COM 기반 (완료)

## 상태: ✅ 완료

## 산출물

### com-core
- ComRuntime: STA 초기화/정리 (RAII)
- Variant: COM VARIANT ↔ Rust enum
- DispatchObject: IDispatch late-binding 래퍼 (get/put/call/get_by)

### excel-bridge
- ExcelApp: 실행, Visible, Quit
- Workbooks → Workbook → Worksheet → Range
- 셀 값 읽기/쓰기, 수식, 저장

### hwp-bridge
- HwpDetector: 설치 감지, 보안 모듈 등록
- HwpApp: 실행, 문서 열기/저장, 텍스트 추출/삽입, HAction

### com-cli
- 대화형 REPL (rustyline)
- excel / hwp / raw 명령어
- raw dispatch 탐색 (get/put/call/chain/target)
