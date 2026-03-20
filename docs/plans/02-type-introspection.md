# Milestone 2: Type 탐색 + HWP Cheat Sheet

## 목표

COM 객체의 메서드/프로퍼티를 자동으로 열거하는 기능 구현.
Office 계열(Excel/Word/PPT)은 LLM 학습 데이터가 풍부하므로 cheat sheet 불필요.
**HWP만** cheat sheet 작성 (학습 데이터 부족).

## 핵심 작업

### 2-1. ITypeInfo 기반 메서드 열거

com-core에 `DispatchObject::list_members()` 추가:
- `IDispatch::GetTypeInfo()` → `ITypeInfo`
- `ITypeInfo::GetTypeAttr()` → 함수/변수 개수
- `ITypeInfo::GetFuncDesc(i)` → 각 함수의 이름, 파라미터, 반환타입, 종류(method/get/put)
- `ITypeInfo::GetVarDesc(i)` → 각 변수(프로퍼티)

출력 형식:
```
[method]  Calculate()
[get]     Name → String
[get/put] Value → Variant
[get]     Range(addr: String) → Range
[get]     Cells(row: I4, col: I4) → Range
```

### 2-2. com-cli에 `raw list` 명령 추가

```
com> raw target excel
com> raw list                    # Application 객체의 전체 멤버
com> raw list methods            # 메서드만
com> raw list props              # 프로퍼티만
```

### 2-3. HWP Cheat Sheet 작성

기존 `docs/analysis/hwp-com-api-reference.md`를 기반으로
LLM 프롬프트에 주입할 수 있는 간결한 형태로 정리.

대상 객체:
- HwpObject (메인)
- HAction / HParameterSet 패턴
- 주요 HParameterSet 이름 + 파라미터

산출물: `docs/cheatsheets/hwp.md`

## 의존성

- Milestone 1 (완료)
- windows crate: `Win32_System_Ole` (ITypeInfo 이미 포함)

## 검증

1. `com-cli`에서 `raw list` 실행 → Excel/HWP 멤버 목록 출력
2. HWP cheat sheet로 LLM이 정확한 HWP 코드를 생성하는지 수동 확인
