# 단계 5: 글자 서식 (CharShape)

## 1. 공식 API 문서 조사

### 서식 적용: HAction + HParameterSet.HCharShape

```js
// 텍스트 선택 후
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Bold = 1;
hwp.HParameterSet.HCharShape.Height = 2400;  // 24pt (100단위)
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);
```

### 서식 읽기: hwp.CharShape (루트 프로퍼티)

```js
// 커서 위치의 서식 읽기
var cs = hwp.CharShape;  // ParameterSet 반환
cs.Item("Height");       // 2400
cs.Item("Bold");         // 1
```

### 주요 프로퍼티

| 프로퍼티 | 타입 | 설명 | 값 예시 |
|----------|------|------|---------|
| Height | I4 | 글자 크기 (100단위, 1pt=100) | 1000=10pt, 2400=24pt |
| Bold | VT(18) | 볼드 | 0/1 |
| Italic | VT(18) | 이탤릭 | 0/1 |
| TextColor | VT(19) | 글자색 (BGR, RGBColor 사용) | RGBColor(255,0,0)=255 |
| UnderlineType | VT(18) | 밑줄 | 0=없음, 1=실선 |
| UnderlineColor | VT(19) | 밑줄 색 | |
| UnderlineShape | VT(18) | 밑줄 모양 | |
| StrikeOutType | VT(18) | 취소선 | 0=없음, 1=실선 |
| StrikeOutColor | VT(19) | 취소선 색 | |
| FaceNameHangul | String | 한글 글꼴명 | |
| FaceNameLatin | String | 영문 글꼴명 | |
| SuperScript | VT(18) | 위첨자 | 0/1 |
| SubScript | VT(18) | 아래첨자 | 0/1 |
| SmallCaps | VT(18) | 작은 대문자 | 0/1 |
| Emboss | VT(18) | 양각 | 0/1 |
| Engrave | VT(18) | 음각 | 0/1 |
| ShadowType | VT(18) | 그림자 종류 | |
| ShadeColor | VT(19) | 음영 색 | |

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 / 예상 | 실제 | 비고 |
|------|-----------------|------|------|
| Underline | `Underline = 1` | **`UnderlineType = 1`** | 프로퍼티명 다름 |
| StrikeOut | `StrikeOut = 1` | **`StrikeOutType = 1`** | 프로퍼티명 다름 |
| 서식 읽기 | HParameterSet + GetDefault | **hwp.CharShape.Item()** | GetDefault는 현재 서식을 안 읽음 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)

### 서식 적용

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 5-1 | Height=2400 (24pt) | ✅ | 글자 크기 변경 |
| 5-2 | Bold=1 | ✅ | 볼드 적용 |
| 5-3 | Italic=1 | ✅ | 이탤릭 적용 |
| 5-4 | TextColor=RGBColor(255,0,0) | ✅ | 빨간색 적용. RGBColor 값=255 (BGR) |
| 5-5 | UnderlineType=1 | ✅ | 밑줄 적용 |
| 5-6 | StrikeOutType=1 | ✅ | 취소선 적용 |
| 5-7 | 서식 설정 후 InsertText | ✅ | 이후 입력에 서식 반영 |
| 5-8 | CreateAction 패턴 | ✅ | SetItem("Height", 1600) 작동 |
| 5-9 | SelectText + CharShape | ✅ | pos 인코딩 + 서식 적용 정상 |

### 서식 읽기

| # | 방법 | 결과 | 비고 |
|---|------|:----:|------|
| 5-10 | HParameterSet + GetDefault | ❌ | 전부 0 반환 — 현재 서식을 안 읽음 |
| 5-11 | **hwp.CharShape.Item()** | ✅ | Height=2400, Bold=1 등 정확히 반환 |

---

## 4. 분석

### 핵심 발견 1: 서식 적용은 두 패턴 모두 가능

```js
// 패턴 A: HParameterSet
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Bold = 1;
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);

// 패턴 B: CreateAction
var act = hwp.CreateAction("CharShape");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Height", 2400);
act.Execute(set);
```

### 핵심 발견 2: 서식 읽기는 hwp.CharShape 사용

`HParameterSet.HCharShape`에 `GetDefault`하면 **기본값(0)**이 채워진다. 현재 커서 위치의 서식을 읽으려면:

```js
var cs = hwp.CharShape;        // 루트 프로퍼티 — 현재 커서의 서식
cs.Item("Height");             // 2400
cs.Item("Bold");               // 1
cs.Item("UnderlineType");      // 1
```

### 핵심 발견 3: 프로퍼티명 주의

| 잘못된 이름 | 올바른 이름 |
|------------|------------|
| `Underline` | `UnderlineType` |
| `StrikeOut` | `StrikeOutType` |

`Underline = 1` 대신 `UnderlineType = 1` 사용.

---

## 5. 최종 결론

### 정상 작동 (✅)

- **Height** — 글자 크기 (100단위)
- **Bold / Italic** — 볼드/이탤릭
- **TextColor** — 글자색 (RGBColor 사용, BGR 순서)
- **UnderlineType** — 밑줄
- **StrikeOutType** — 취소선
- **서식 먼저 설정 → InsertText** — 이후 입력에 서식 반영
- **hwp.CharShape.Item()** — 현재 커서 서식 읽기
- **SelectText + CharShape** — pos 인코딩으로 범위 서식 적용

### 주의 필요 (⚠️)

- **서식 읽기**: `HParameterSet + GetDefault` 아닌 `hwp.CharShape.Item()` 사용
- **밑줄/취소선**: `Underline`이 아니라 `UnderlineType`, `StrikeOut`이 아니라 `StrikeOutType`
