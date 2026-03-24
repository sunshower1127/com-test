# 단계 1: 문서 생명주기

## 1. 공식 API 문서 조사

### 출처

- [한컴 개발자 포럼 - SaveAs Format](https://forum.developer.hancom.com/t/saveas/660)
- [한컴 개발자 포럼 - Clear option](https://forum.developer.hancom.com/t/topic/1424)
- [한컴 개발자 포럼 - API Lifecycle](https://forum.developer.hancom.com/t/api/1582)
- [pyhwpx.com - Open/Close](https://pyhwpx.com/272)

### Open(filename, Format, arg)

- **filename**: 파일 전체 경로
- **Format**: 대문자. "HWP", "HWPX", "TEXT", "HTML", "OOXML" 등. 생략 시 확장자로 자동 감지
- **arg**: 세미콜론 구분 옵션
  - `forceopen:true` — 잠긴 파일 강제 열기
  - `suspendpassword:true` — 암호 팝업 억제
  - `versionworning:false` — 버전 경고 억제 (오타 아님, API가 이렇게 씀)
- **반환**: Bool

### SaveAs(Path, Format, arg)

- **Path**: 저장할 전체 경로
- **Format**: 대문자 필수. 지원 형식:

| Format | 설명 |
|--------|------|
| `"HWP"` | HWP 5 (기본) |
| `"HWPX"` | HWPX (OWPML) |
| `"PDF"` | PDF |
| `"OOXML"` | MS Word DOCX |
| `"DOCRTF"` | MS Word DOC |
| `"RTF"` | Rich Text Format |
| `"HTML"` | HTML (데이터 중심) |
| `"HTML+"` | HTML (레이아웃 보존) |
| `"TEXT"` | ASCII 텍스트 |
| `"UNICODE"` | 유니코드 텍스트 |
| `"HWPML2X"` | HWPML 2.x (.hml) |
| `"HWP30"` / `"HWP20"` | 레거시 형식 |

- **arg**: `lock:false` 등
- **반환**: Bool

### Save(save_if_dirty)

- `True` 또는 생략 — 수정된 경우만 저장
- `False` — 무조건 저장

### Clear(option)

| option | 동작 |
|--------|------|
| 0 | 수정 시 저장 대화상자 표시 |
| 1 | 저장 안 하고 닫기 (팝업 없음) |
| 2 | 수정 시 저장 후 닫기 |
| 3 | 무조건 저장 후 닫기 |

### Quit()

파라미터 없음. HWP 프로세스 종료.

### Run("FileNew")

새 빈 문서 생성.

### XHwpDocuments

- `.Add(isTab)` — true=탭, false=새 창
- `.Count` — 열린 문서 수
- `.Close(isDirty)` — 현재 문서 닫기
- `.Active_XHwpDocument` — 활성 문서

### Insert(Path, Format, arg)

커서 위치에 파일 내용 삽입. Open과 동일한 파라미터 구조.

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 | 실험 결과 | 비고 |
|------|-----------|-----------|------|
| Open 보안 팝업 | 보안모듈 필요 언급 | 팝업은 뜨지만 블로킹 → 허용 시 성공 | 보안모듈 없이 사용 가능 |
| SaveAs 보안 팝업 | 보안모듈 필요 언급 | 논블로킹으로 즉시 에러 (0x80010105) | **보안모듈 필수** |
| 경로 구분자 | 백슬래시 사용 | 슬래시, 백슬래시 모두 작동 | JS에서는 `/` 권장 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)
- 보안모듈 미설치 상태

### Open

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 1-1 | Open (슬래시 경로) | ✅ | 팝업 허용 후 성공, 블로킹 |
| 1-2 | Open (백슬래시 경로) | ✅ | 동일 |
| 1-3 | Open 후 Path | ✅ | 정확한 경로 반환 |
| 1-4 | Open 후 IsEmpty | ✅ | false |
| 1-5 | Open 후 GetTextFile | ✅ | 내용 정상 추출 |
| 1-6 | Open 반환값 | ✅ | true |
| 1-7 | "모두 허용" 후 재 Open | ✅ | 팝업 없이 바로 성공 |

### SaveAs

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 1-8 | SaveAs HWP | ❌ | 0x80010105 — 보안모듈 필요 |
| 1-9 | SaveAs HWPX | ❌ | 동일 |
| 1-10 | SaveAs PDF | ❌ | 동일 |
| 1-11 | SaveAs TEXT | ❌ | 동일 |
| 1-12 | SaveAs HTML | ❌ | 동일 |
| 1-13 | SaveAs DOCX | ❌ | 동일 |

### Save

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 1-14 | Save(false) | ⚠️ | 에러는 안 나지만 IsModified가 여전히 true — 실질적 저장 안 됨 |

### Clear / Quit

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 1-15 | Clear(1) | ✅ | 저장 안 하고 닫기, IsEmpty=true |
| 1-16 | Quit() | ✅ | 정상 종료 |

### FileNew / XHwpDocuments

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 1-17 | Run("FileNew") | ✅ | 새 빈 문서 생성 |
| 1-18 | XHwpDocuments.Count | ✅ | 정확한 문서 수 반환 |
| 1-19 | XHwpDocuments.Add(false) | ✅ | Count 1→2 |
| 1-20 | XHwpDocuments.Close(false) | ✅ | Count 2→1 |

---

## 4. 분석

### 핵심 발견: Open vs SaveAs 보안 동작 차이

| | Open | SaveAs |
|--|------|--------|
| 보안모듈 없이 | ✅ 작동 | ❌ 실패 |
| 팝업 동작 | **블로킹** — 허용 누를 때까지 대기 | **논블로킹** — 즉시 에러 반환 |
| "모두 허용" 효과 | 세션 동안 이후 팝업 안 뜸 | 효과 없음 (이미 에러) |
| 에러 코드 | 없음 (성공) | 0x80010105 (RPC_E_SERVERFAULT) |

### 경로 형식

슬래시(`/`), 백슬래시(`\`) 모두 작동. HWP COM이 내부적으로 정규화.

---

## 5. 최종 결론

### 정상 작동 (✅)

- **Open** — 보안모듈 없이 작동 (팝업 허용 필요, 블로킹)
- **Clear(1)** — 저장 안 하고 닫기
- **Quit()** — 정상 종료
- **Run("FileNew")** — 새 문서
- **XHwpDocuments.Add/Close/Count** — 문서 관리
- **경로**: 슬래시, 백슬래시 모두 가능

### 보안모듈 필요 (❌)

- **SaveAs** — 모든 형식 (HWP, HWPX, PDF, TEXT, HTML, DOCX)
- **Save** — 실질적 저장 안 됨

### 미테스트 (보안모듈 설치 후)

- SaveAs 각 포맷별 실제 저장 확인
- Save 동작 확인
- Insert (파일 삽입) — 보안 팝업 동작 미확인

---

## TODO: 보안모듈 설치

SaveAs/Save 테스트를 위해 보안모듈 설치가 필요함.

### 1. DLL 다운로드

한컴 공식 GitHub에서 다운로드:
```
https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/보안모듈(Automation).zip
```

ZIP 내용:
- `FilePathCheckerModuleExample.dll`
- 레지스트리 설정 스크린샷
- 소스 코드

### 2. DLL 배치

원하는 위치에 DLL 복사. 예시:
```
C:\Program Files (x86)\HNC\Security\FilePathCheckerModuleExample.dll
```

### 3. 레지스트리 등록

`regedit`에서:
```
HKEY_CURRENT_USER\SOFTWARE\HNC\HwpAutomation\Modules
```
- 키가 없으면 새로 생성
- 문자열 값 추가:
  - 이름: `FilePathCheckerModuleExample`
  - 값: DLL 전체 경로 (**따옴표 없이!**)

### 4. 코드에서 호출

```js
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
```

- 두 번째 인자는 레지스트리 값 이름과 **정확히 일치**해야 함
- RegisterModule이 `false`를 반환해도 작동하는 경우 있음 (알려진 버그)

### 5. 버전별 참고

| 버전 | 상태 |
|------|------|
| 한글 2018 | RegisterModule 정상 |
| 한글 2020 | 정상 |
| 한글 2022 | 정상 (ActiveX 마지막 지원) |
| 한글 2024 | RegisterModule 버그 보고 있음 — 팝업 여전히 뜰 수 있음 |

### 6. 설치 후 확인 방법

```js
// 보안모듈 등록
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
// SaveAs 시도 — 팝업 없이 성공하면 OK
hwp.SaveAs("C:/test.hwp", "HWP", "");
```
