# 단계 9: 필드

## 1. 공식 API 문서 조사

### CreateField(Direction, memo, name)

- **Direction**: 필드 종류 ("SET", "GET", "CLICK" 등)
- **memo**: 메모 (빈 문자열 가능)
- **name**: 필드 이름
- **반환**: Bool

### PutFieldText(name, text) / GetFieldText(name)

- 필드 이름으로 값 설정/읽기

### GetFieldList(option1, option2)

- 문서의 필드 목록 반환 (`\x02` 구분자)

### FieldExist(name)

- 필드 존재 여부 (Bool)

### MoveToField(name, text, start, select)

- 필드로 커서 이동

---

## 2. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)

### 필드 생성

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 9-1 | CreateField("SET", "", "name") | ✅ | ret=true |
| 9-2 | CreateField 여러 개 | ✅ | 모두 ret=true |
| 9-3 | Direction: SET/GET/CLICK/REVISION 등 | ✅ | 전부 생성 가능 |

### 필드 목록/존재

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 9-4 | GetFieldList(0, 0) | ✅ | `\x02` (charCode 2) 구분자로 모든 필드명 반환 |
| 9-5 | FieldExist | ✅ | 모든 필드에 true 반환 |

### 필드 값 설정/읽기

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 9-6 | PutFieldText (첫 번째 필드) | ✅ | 값 설정 정상 |
| 9-7 | GetFieldText (첫 번째 필드) | ✅ | 값 읽기 정상 |
| 9-8 | PutFieldText (두 번째 이후) | ❌ | **빈 문자열 — 값 설정 안 됨** |
| 9-9 | GetFieldText (두 번째 이후) | ❌ | **빈 문자열 반환** |
| 9-10 | PutFieldText 구분자로 한번에 | ❌ | 첫 필드만 설정됨 |

### 필드 이동

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 9-11 | MoveToField (첫 번째 필드) | ✅ | ret=true |
| 9-12 | MoveToField (두 번째 이후) | ❌ | ret=false |

---

## 3. 분석

### 핵심 발견: 첫 번째 필드만 접근 가능

CreateField로 여러 필드를 만들 수 있고, GetFieldList/FieldExist에서도 전부 보이지만, **PutFieldText/GetFieldText/MoveToField는 첫 번째 필드에서만 작동**.

이 제한은 다음과 무관:
- Direction 값 (SET/GET 등)
- 줄바꿈 방식 (BreakPara, `\r\n`)
- 같은 줄/다른 줄
- 구분자 한번에/개별 호출

**한글 2018 HwpObject의 필드 접근 제한으로 추정.** 다른 버전에서는 다를 수 있음.

### GetFieldList 구분자

```js
var list = hwp.GetFieldList(0, 0);
// "field1\x02field2\x02field3" (\x02 = STX, charCode 2)
var fields = list.split('\x02');
```

---

## 4. 최종 결론

### 정상 작동 (✅)

- **CreateField** — 필드 생성 (여러 종류)
- **GetFieldList** — 전체 필드 목록 (`\x02` 구분)
- **FieldExist** — 존재 확인
- **PutFieldText / GetFieldText** — 첫 번째 필드만
- **MoveToField** — 첫 번째 필드만

### 작동하지 않음 (❌)

- **두 번째 이후 필드** — PutFieldText/GetFieldText/MoveToField 전부 안 됨
- 필드가 1개면 완벽하게 작동, 2개 이상이면 첫 번째만 접근 가능
