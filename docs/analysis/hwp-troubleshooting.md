# HWP COM 자동화 트러블슈팅 & 실전 노하우

## 32비트 / 64비트 호환성

### 현재 환경 (이 프로젝트)

- HWP 2018: **32비트** (`C:\Program Files (x86)\Hnc\Office 2018`)
- Rust 바이너리: **64비트** (`x86_64-pc-windows-msvc`)
- 결과: **정상 동작** — LocalServer32 (out-of-process COM)이므로 OS가 자동 마샬링

### 일반 규칙

- HWP COM은 `LocalServer32` 등록 → 별도 프로세스로 실행됨
- 32↔64 비트 간 IPC는 COM 런타임이 자동 처리
- Python에서는 비트를 맞추라는 조언이 많지만, out-of-process COM에서는 불필요
- **InprocServer32 (in-process)였다면** 반드시 비트를 맞춰야 함

## 프로세스 종료 / 좀비 방지

### 올바른 종료 순서

```
1. hwp.Clear(1)    — 문서 닫기 (1 = 저장 안 함)
2. hwp.Quit()      — HWP 프로세스 종료
3. COM 객체 해제    — Rust: Drop trait / Python: del hwp
```

### 주의사항

- `Visible = False`로 숨겨도 `Quit()` 안 하면 **백그라운드에 프로세스 남음**
- `FileClose`만으로는 프로세스가 종료되지 않음
- 예외 발생 시에도 Quit이 호출되도록 **Drop trait / try-finally 필수**
- 최후의 안전장치: `taskkill /f /im Hwp.exe`

### Rust에서의 구현

```rust
impl Drop for HwpApp {
    fn drop(&mut self) {
        let _ = self.clear();   // Clear(1)
        let _ = self.quit();    // Quit
        // IDispatch Drop이 Release를 호출
    }
}
```

## 한컴 버전별 호환성

| 항목 | 2018/2020 | 2022 | 2024 |
|------|-----------|------|------|
| ActiveX (HwpCtrl) | O | O | **X (제거됨)** |
| HwpAutomation | O | O | O |
| 다중 인스턴스 | O | O | **마지막 1개만** |
| Visible 제어 | 정상 | 정상 | **작동 안 할 수 있음** |
| 레지스트리 경로 | HwpCtrl or HwpAutomation | 둘 다 | HwpAutomation만 |
| 안정성 | 안정 | 안정 | 패치 후 오류 보고 다수 |

### 2024 주요 변경사항

- **HwpCtrl (ActiveX) 제거** → HwpAutomation으로 전환 필수
- **다중 인스턴스 불가** → 마지막 창 1개만 열림 (자동화에 큰 영향)
- Visible 제어가 안 되는 사례 보고
- RegisterModule 후에도 팝업 지속되는 버그
- `hwp.exe -Regserver`로 COM 재등록 시 해결되는 경우 있음

## 자주 겪는 에러

### 1. "보안 정책상 사용할 수 없는 기능입니다"

- 원인: 보안 모듈 미등록
- 해결: DLL 다운로드 → 레지스트리 등록 → RegisterModule 호출

### 2. "매개 변수의 개수가 잘못되었습니다" (0x8002000E)

- 원인: COM 메서드에 필수 인자 누락
- 해결: HWP API는 선택 인자를 잘 지원 안 함 → 빈 문자열이라도 전달
- 예: `SaveAs(path, format, "")` — 3번째 arg 필수

### 3. "서버에서 예외 오류가 발생했습니다" (0x80010105)

- 원인: HWP 내부에서 예외 발생 (보안 차단, 잘못된 파일 경로 등)
- 해결: 보안 모듈 확인, 절대 경로 사용, `forceopen:true` 옵션

### 4. GetTextFile 한글 깨짐

- 원인: `GetTextFile("TEXT", "")` → ANSI (CP949) 반환
- **해결**: `GetTextFile("UNICODE", "")` 사용 필수

### 5. Python EnsureDispatch 오류

- 원인: gen_py 캐시 꼬임
- 해결: 캐시 폴더 삭제, 또는 `Dispatch()`(late binding) 사용
- Rust는 IDispatch late binding이므로 이 문제 없음

### 6. 파일 경로 오류

- HWP COM은 **절대 경로만 지원**
- 상대 경로 → 파일 못 찾음
- UNC prefix (`\\?\`) 제거 필요

## Rust/Go/C++ 에서 HWP COM 사용 현황

| 언어 | COM 자동화 | 파일 파싱 |
|------|-----------|----------|
| **Rust** | 사례 없음 **(이 프로젝트가 최초)** | hwp-rs, hwpers, unhwp |
| **Go** | go-ole 기반 gohwp 패키지 존재 | — |
| **C++** | CoCreateInstance + IDispatch 사례 있음 | — |
| **C#** | COM interop 사례 다수, 버전 간 호환 문제 보고 | — |
| **Python** | pyhwpx, hwpapi 등 성숙한 래퍼 다수 | pyhwp |

### 참고 프로젝트

- **hwp-rs** (Rust 파서): https://github.com/hahnlee/hwp-rs
- **hwpers** (Rust 파서): https://crates.io/crates/hwpers
- **gohwp** (Go COM): https://www.byul.pe.kr/2024/04/04-gohwp-gohwp-package.html
- **pyhwpx** (Python 래퍼): https://github.com/martiniifun/pyhwpx
- **hwpapi** (Python 래퍼): https://github.com/JunDamin/hwpapi

## 출처

- https://forum.developer.hancom.com/ (한컴디벨로퍼 포럼)
- https://pyhwpx.com/ (pyhwpx 공식 블로그)
- https://wikidocs.net/book/8956 (pyhwpx Cookbook)
- https://www.byul.pe.kr/2024/04/01-gohwp-go-ole-package.html (Go HWP 자동화)
