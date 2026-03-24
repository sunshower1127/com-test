# 단계 3: 커서 이동 + 선택

## 1. 공식 API 문서 조사

### 출처

- [한컴 개발자 - MovePos](https://developer.hancom.com/en-us/webhwp/devguide/hwpctrl/methods/movepos)
- [한컴 개발자 - SetPos](https://developer.hancom.com/webhwp/devguide/hwpctrl/methods/setpos)
- [pyhwpx - 커서 이동](https://wikidocs.net/257898)
- [martinii.fun - Run 전체 목록](https://martinii.fun/pages/hwp%EC%9D%98-Run%EB%A9%94%EC%84%9C%EB%93%9C-%EC%A0%84%EC%B2%B4%EB%AA%A9%EB%A1%9D-%EB%B0%8F-%EC%8B%9C%EC%97%B0%ED%99%94%EB%A9%B4)

### Run 기반 커서 이동

파라미터 없는 단순 액션. `hwp.Run("ActionName")` 으로 호출.

| 액션 | 설명 |
|------|------|
| `MoveDocBegin` / `MoveDocEnd` | 문서 처음/끝 |
| `MoveLineBegin` / `MoveLineEnd` | 줄 처음/끝 |
| `MoveParaBegin` / `MoveParaEnd` | 문단 처음/끝 |
| `MoveNextParaBegin` / `MovePrevParaBegin` | 다음/이전 문단 처음 |
| `MoveLeft` / `MoveRight` | 한 글자 좌/우 |
| `MoveUp` / `MoveDown` | 한 줄 위/아래 |
| `MoveNextWord` / `MovePrevWord` | 다음/이전 단어 |
| `MovePageUp` / `MovePageDown` | 페이지 위/아래 |

### Run 기반 선택 (MoveSel*)

이동 액션에 `Sel` 접두어를 붙이면 선택하면서 이동. `MoveSelRight`, `MoveSelDocEnd` 등.

### MovePos(moveID, para, pos)

- **moveID**: 이동 종류 (2=문서 처음, 3=문서 끝, 6=문단 처음, 7=문단 끝 등)
- **para**: 문단 번호 (moveID 0,1일 때만 사용)
- **pos**: 문단 내 오프셋 (**글자 인덱스가 아님** — 아래 pos 인코딩 참고)

### SetPos(list, para, pos)

- **list**: 문서 리스트 ID (0=본문)
- **para**: 문단 번호 (0-based)
- **pos**: 문단 내 오프셋 (**글자 인덱스가 아님**)

### GetPosBySet() / SetPosBySet(set)

- `GetPosBySet()`: 현재 위치를 ParameterSet(SetID="ListParaPos")으로 반환
- `SetPosBySet(set)`: 저장된 위치로 이동
- `Item("List")`, `Item("Para")`, `Item("Pos")` 로 값 읽기 가능

### SelectText(spara, spos, epara, epos)

- 프로그래밍으로 텍스트 영역 선택
- **spos/epos는 pos 인코딩 사용** (글자 인덱스가 아님)
- 올바른 pos를 넣으면 saveblock, Delete, InsertText 모두 정상 작동

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 | 실험 결과 | 비고 |
|------|-----------|-----------|------|
| pos 파라미터 | 글자 인덱스 | **HWP 내부 오프셋** | 첫 문단은 16+, 이후 문단은 0+ |
| GetPosBySet Item | 값 반환 | 브릿지에서 null이었음 | **VT_VARIANT 처리 추가로 해결** |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)
- 테스트 텍스트: 영문 (자동교정 회피)

### Run 기반 커서 이동

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 3-1 | MoveDocBegin | ✅ | 문서 처음으로 정확히 이동 |
| 3-2 | MoveDocEnd | ✅ | 문서 끝으로 정확히 이동 |
| 3-3 | MoveLineBegin | ✅ | 줄 처음 |
| 3-4 | MoveLineEnd | ✅ | 줄 끝 |
| 3-5 | MoveNextParaBegin | ✅ | 다음 문단 처음 |
| 3-6 | MovePrevParaBegin | ✅ | 이전 문단 처음 |
| 3-7 | MoveParaBegin | ✅ | 현재 문단 처음 |
| 3-8 | MoveParaEnd | ✅ | 현재 문단 끝 |

### Run 기반 선택

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 3-9 | MoveSelDocEnd | ✅ | 전체 텍스트 선택 |
| 3-10 | MoveSelLineEnd | ✅ | 첫 줄 전체 선택 |
| 3-11 | MoveSelNextWord | ✅ | 다음 단어 선택 |
| 3-12 | MoveSelRight x3 | ✅ | 정확히 3글자 선택 |

### MovePos / SetPos (pos 인코딩 적용 후)

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 3-13 | MovePos(2) = 문서 처음 | ✅ | moveTopOfFile |
| 3-14 | MovePos(3) = 문서 끝 | ✅ | moveBottomOfFile |
| 3-15 | MovePos(1,1,0) = 2번째 문단 처음 | ✅ | para 작동, pos=0 (이후 문단) |
| 3-16 | MovePos(1,0,16+3) = 첫문단 4번째 | ✅ | **pos 인코딩 적용 시 정상** |
| 3-17 | MovePos(6) = 현재 문단 처음 | ✅ | moveStartOfPara |
| 3-18 | MovePos(7) = 현재 문단 끝 | ✅ | moveEndOfPara |

### SetPos / GetPosBySet

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 3-19 | SetPos(0,0,16+N) 첫 문단 | ✅ | 영문/한글/특수문자 모두 글자당 1 |
| 3-20 | SetPos(0,1,N) 이후 문단 | ✅ | 오프셋 0부터 |
| 3-21 | SetPos 100글자 긴 텍스트 | ✅ | 정상 |
| 3-22 | GetPosBySet → Item("Pos") | ✅ | **VT_VARIANT 처리 후 정상** |
| 3-23 | GetPosBySet → SetPosBySet | ✅ | 위치 저장/복원 정상 |

### SelectText (pos 인코딩 적용 후)

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 3-24 | SelectText(0,16+2,0,16+5) → saveblock | ✅ | "CDE" 정상 선택 |
| 3-25 | SelectText + Delete | ✅ | 선택 영역 정상 삭제 |
| 3-26 | SelectText 여러 문단 | ✅ | para별 올바른 pos 적용 시 정상 |

### Undo

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 3-27 | Run("Undo") | ✅ | Delete 후 Undo 정상 복원 |

---

## 4. 분석

### 핵심 발견 1: pos는 글자 인덱스가 아니라 HWP 내부 오프셋

공식 문서에서 pos를 "글자 인덱스"처럼 설명하지만, 실제로는 **문단 데이터의 바이트 오프셋**이다.

| 문단 | pos 공식 | 이유 |
|------|---------|------|
| **첫 문단 (para=0)** | `기본 오프셋 + 글자인덱스` | 섹션 속성/칼럼 속성 등 컨트롤 문자가 앞에 존재 |
| **이후 문단 (para>0)** | `글자인덱스` | 컨트롤 문자 없음 |

기본 문서에서 첫 문단 오프셋은 16이지만, 머리말/꼬리말 등 컨트롤이 추가되면 달라질 수 있다.

**HWPX 분석으로 확인된 원인:**

```xml
<!-- 첫 문단: secPr + colPr 컨트롤이 텍스트 앞에 존재 -->
<hp:p id="0">
  <hp:run><hp:secPr .../><hp:ctrl><hp:colPr .../></hp:ctrl></hp:run>
  <hp:run><hp:t>ABC</hp:t></hp:run>
</hp:p>

<!-- 이후 문단: 바로 텍스트 -->
<hp:p id="2147483648">
  <hp:run><hp:t>DEF</hp:t></hp:run>
</hp:p>
```

**안전한 pos 계산 방법:**

```js
// 하드코딩 대신 동적으로 오프셋 읽기
function getParaOffset(para) {
  hwp.SetPos(0, para, 0);
  hwp.Run("MoveParaBegin");
  var ps = hwp.GetPosBySet();
  return ps.Item("Pos");  // 첫 문단: 16 (보통), 이후: 0
}

// 사용
var offset = getParaOffset(0);
hwp.SetPos(0, 0, offset + 5);  // 첫 문단 6번째 글자
```

### 핵심 발견 2: GetPosBySet Item 읽기 — VT_VARIANT 처리로 해결

`GetPosBySet().Item("Pos")` 가 null을 반환하던 문제는 **브릿지의 VT_VARIANT (중첩 Variant) 미처리** 때문이었다.

HWP ParameterSet의 `Item()` 메서드는 `VT_VARIANT` 안에 `VT_I4`를 감싸서 반환한다. `variant_from_raw`에 `VT_VARIANT` → 재귀 언래핑을 추가하여 해결.

```
수정 전: Item("Pos") → VT_VARIANT(VT_I4(16)) → _ 매치 → Empty → null
수정 후: Item("Pos") → VT_VARIANT(VT_I4(16)) → 재귀 → I32(16) → 16
```

### 핵심 발견 3: SelectText()는 정상 작동한다

이전에 "가짜 선택"이라고 결론 내렸지만, **pos 인코딩이 잘못됐을 뿐**이었다.

| | 잘못된 pos (글자 인덱스) | 올바른 pos (내부 오프셋) |
|--|:----------------------:|:----------------------:|
| saveblock | ❌ 빈 문자열 | ✅ 정상 |
| Delete | ❌ 무시됨 | ✅ 정상 |
| InsertText 덮어쓰기 | ❌ 앞에 삽입 | ✅ 정상 |

```js
// ❌ 안 됨 — 글자 인덱스를 넣으면 실패
hwp.SelectText(0, 2, 0, 5);

// ✅ 됨 — 올바른 pos 인코딩
var offset = getParaOffset(0);
hwp.SelectText(0, offset+2, 0, offset+5);  // "CDE" 선택
hwp.Run("Delete");  // 정상 삭제
```

---

## 5. 최종 결론

### 정상 작동 (✅)

- **Run 기반 커서 이동** — 전부 정상
- **Run 기반 선택** — MoveSelRight 등 + saveblock/Delete 조합 정상
- **MovePos** — moveID별 이동 정상. pos는 내부 오프셋 사용
- **SetPos** — para + pos (내부 오프셋) 정상
- **SelectText** — pos 인코딩 적용 시 saveblock/Delete/InsertText 모두 정상
- **GetPosBySet** — Item("List"/"Para"/"Pos") 읽기 정상 (VT_VARIANT 처리 후)
- **SetPosBySet** — 위치 복원 정상
- **Undo** — 정상

### 주의 필요 (⚠️)

- **pos는 글자 인덱스가 아님** — 첫 문단은 오프셋 존재 (기본 16, 문서 구조에 따라 변동 가능). `GetPosBySet().Item("Pos")`로 동적 감지 권장
- **영문/한글/특수문자 모두 글자당 pos 1** — 멀티바이트 걱정 불필요
