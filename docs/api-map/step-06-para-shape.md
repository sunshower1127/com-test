# 단계 6: 문단 서식 (ParagraphShape)

## 1. 공식 API 문서 조사

### 서식 적용: HAction + HParameterSet.HParaShape

```js
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet);
hwp.HParameterSet.HParaShape.AlignType = 3;  // 가운데
hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet);
```

### 서식 읽기: hwp.ParaShape (루트 프로퍼티)

```js
var ps = hwp.ParaShape;
ps.Item("AlignType");    // 3 = 가운데
ps.Item("LeftMargin");   // 2000
```

### 주요 프로퍼티

| 프로퍼티 | 설명 | 값 예시 |
|----------|------|---------|
| AlignType | 정렬 | 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데 |
| LeftMargin | 왼쪽 여백 (HWP 단위) | 2000 ≈ 7mm |
| RightMargin | 오른쪽 여백 | |
| Indentation | 첫 줄 들여쓰기 | 1000 ≈ 3.5mm |
| LineSpacing | 줄간격 값 | ⚠️ 조사 필요 |
| LineSpacingType | 줄간격 종류 | 0=%, 1=고정, 2=여백만 등 |
| PrevSpacing | 문단 앞 간격 | 500 |
| NextSpacing | 문단 뒤 간격 | 500 |
| KeepWithNext | 다음 문단과 함께 | 0/1 |
| KeepLinesTogether | 문단 나누지 않음 | 0/1 |
| PagebreakBefore | 문단 앞에서 페이지 나눔 | 0/1 |
| BreakLatinWord | 영어 단어 잘림 | |
| TextAlignment | 세로 정렬 | 0=기본 |
| TabDef | 탭 정의 | |
| BorderFill | 테두리/배경 | |

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 / 예상 | 실제 | 비고 |
|------|-----------------|------|------|
| 정렬 | `Alignment` | **`AlignType`** | 프로퍼티명 다름 |
| 문단 앞 간격 | `SpaceBeforePara` | **`PrevSpacing`** | 프로퍼티명 다름 |
| 문단 뒤 간격 | `SpaceAfterPara` | **`NextSpacing`** | 프로퍼티명 다름 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)

### 서식 적용

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 6-1 | AlignType=3 (가운데) | ✅ | |
| 6-2 | AlignType=2 (오른쪽) | ✅ | |
| 6-3 | LeftMargin=2000 | ✅ | 왼쪽 들여쓰기 |
| 6-4 | PrevSpacing/NextSpacing=500 | ✅ | 문단 앞뒤 간격 |
| 6-5 | LineSpacing=250 (단독) | ⚠️ | GetDefault가 리셋 → 160. **LineSpacingType과 함께 설정 시 정상** |
| 6-6 | Indentation=1000 | ✅ | 첫 줄 들여쓰기 |
| 6-7 | CreateAction 패턴 | ✅ | SetItem("AlignType", 3) 작동 |

### 서식 읽기

| # | 읽기 | 결과 | 비고 |
|---|------|:----:|------|
| 6-8 | AlignType | ✅ | 3 (가운데), 2 (오른쪽) 정확 |
| 6-9 | LeftMargin | ✅ | 2000 정확 |
| 6-10 | PrevSpacing / NextSpacing | ✅ | 500/500 정확 |
| 6-11 | LineSpacing | ⚠️ | 250 설정 → 160 반환. 단위 변환 또는 브릿지 이슈 |
| 6-12 | LineSpacingType | ✅ | 0 반환 |

---

## 4. 분석

### 핵심 발견 1: 프로퍼티명이 공식 문서와 다름

CharShape와 마찬가지로 HParaShape도 실제 프로퍼티명이 다르다.

| 잘못된 이름 | 올바른 이름 |
|------------|------------|
| `Alignment` | `AlignType` |
| `SpaceBeforePara` | `PrevSpacing` |
| `SpaceAfterPara` | `NextSpacing` |

### 핵심 발견 2: 서식 읽기는 hwp.ParaShape 사용

CharShape와 동일한 패턴. `HParameterSet.HParaShape` + `GetDefault`가 아니라 **`hwp.ParaShape.Item()`** 으로 현재 커서 위치의 문단 서식을 읽는다.

### 핵심 발견 3: LineSpacing은 LineSpacingType과 반드시 함께 설정

`GetDefault`가 **모든 프로퍼티를 기본값(160, Type=0)으로 리셋**한다. `LineSpacing`만 바꾸면 다른 프로퍼티가 기본값으로 돌아가면서 충돌.

```js
// ❌ LineSpacing만 설정 → 160으로 리셋됨
hwp.HAction.GetDefault("ParagraphShape", ...);
hwp.HParameterSet.HParaShape.LineSpacing = 250;
hwp.HAction.Execute(...);
// → 읽으면 160

// ✅ LineSpacingType과 함께 설정
hwp.HAction.GetDefault("ParagraphShape", ...);
hwp.HParameterSet.HParaShape.LineSpacingType = 0;  // 0=비율%
hwp.HParameterSet.HParaShape.LineSpacing = 250;
hwp.HAction.Execute(...);
// → 읽으면 250 ✅
```

| LineSpacingType | 의미 | LineSpacing 단위 |
|:---------------:|------|-----------------|
| 0 | 비율 (%) | 160 = 160% |
| 1 | 고정값 | HWP 단위 |
| 2 | 여백만 지정 | HWP 단위 |
| 3 | 최소 | HWP 단위 |
| 4 | (미확인) | |

**이 패턴은 다른 프로퍼티에도 적용될 수 있음** — GetDefault가 전체 리셋하므로, 변경하려는 프로퍼티와 관련 프로퍼티를 모두 명시적으로 설정하는 것이 안전.

---

## 5. 최종 결론

### 정상 작동 (✅)

- **AlignType** — 정렬 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데)
- **LeftMargin / RightMargin** — 좌우 여백
- **Indentation** — 첫 줄 들여쓰기
- **PrevSpacing / NextSpacing** — 문단 앞뒤 간격
- **LineSpacing** — 줄간격 (적용은 됨)
- **hwp.ParaShape.Item()** — 서식 읽기
- **CreateAction 패턴** — 정상

### 주의 필요 (⚠️)

- **LineSpacing** — 반드시 `LineSpacingType`과 함께 설정. 단독 설정 시 GetDefault가 리셋
- **GetDefault** — 모든 프로퍼티를 기본값으로 리셋하므로, 변경할 프로퍼티와 관련 프로퍼티를 모두 명시 설정할 것
