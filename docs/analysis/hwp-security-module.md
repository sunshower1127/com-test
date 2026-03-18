# HWP 보안 모듈 (FilePathCheckDLL) 분석

## 개요

한컴이 COM 자동화 API의 파일 시스템 무단 접근을 막기 위해 만든 보안 장치.
보안 모듈 미등록 시 파일 열기/저장마다 **보안 승인 팝업**이 표시되거나 API 호출이 차단됨.

## 설치 절차

### 1. DLL 다운로드

- **한컴디벨로퍼**: https://developer.hancom.com/hwpautomation
- **GitHub 아카이브**: https://github.com/hancom-io/devcenter-archive/raw/main/hwp-automation/보안모듈(Automation).zip

압축 해제 시 포함 파일:
- `FilePathCheckerModuleExample.dll` — 실제 보안 모듈
- 소스 코드 (참고용)

### 2. DLL을 원하는 위치에 배치

예: `C:\tools\FilePathCheckerModuleExample.dll`

### 3. 레지스트리 등록

| 사용 방식 | 레지스트리 경로 |
|-----------|----------------|
| **HwpAutomation** (`hwpframe.hwpobject`) | `HKCU\SOFTWARE\HNC\HwpAutomation\Modules` |
| **HwpCtrl** (ActiveX) | `HKCU\SOFTWARE\HNC\HwpCtrl\Modules` |

> **주의**: HwpAutomation을 쓰면서 HwpCtrl 경로에 등록하면 **동작 안 함**.

등록 방법:
1. `regedit` → 해당 경로 이동
2. 새로 만들기 → 문자열 값
3. **이름**: `FilePathCheckerModule`
4. **값**: DLL 절대 경로 (따옴표 포함하면 안 됨)

### 4. COM 호출

```python
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
# 첫 번째 인자: 항상 고정 "FilePathCheckDLL"
# 두 번째 인자: 레지스트리에 등록한 문자열 값의 "이름"
# 반환값: True(성공) / False(실패)
```

## RegisterModule 실패 원인

| 원인 | 설명 |
|------|------|
| DLL 파일 미존재 | 레지스트리 경로에 실제 DLL이 없음 |
| 이름 불일치 | RegisterModule 두 번째 인자 ≠ 레지스트리 값 이름 |
| 레지스트리 경로 오류 | HwpAutomation인데 HwpCtrl에 등록 (또는 반대) |
| 경로에 따옴표 포함 | `"C:\path\to\dll"` → 따옴표 제거 필요 |
| 한글 2024 버그 | COM 재등록 필요: `hwp.exe -Regserver` |

## 보안 모듈 없이 할 수 있는 것

| 동작 | 보안 모듈 없이 | 보안 모듈 있으면 |
|------|---------------|-----------------|
| COM 객체 생성 | O | O |
| 텍스트 삽입 | O | O |
| 텍스트 추출 (GetTextFile) | O | O |
| 파일 열기 (Open) | 팝업 또는 차단 | O |
| 파일 저장 (SaveAs) | 팝업 또는 차단 | O |

## pyhwpx (Python 래퍼)

`pyhwpx` 패키지는 보안 모듈 DLL을 내장하고 있어 `Hwp()` 생성 시 자동 등록.
Rust에서는 이 접근이 불가능하므로 수동 설치가 필요.

## 출처

- https://pyhwpx.com/67
- https://wikidocs.net/179410
- https://forum.developer.hancom.com/t/registermodule/165
- https://blog.dhlee.info/226
