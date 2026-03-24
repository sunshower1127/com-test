# 단계 4: 텍스트 수정

## 1. 공식 API 문서 조사

### 출처

- [한컴 개발자 - AllReplace](https://developer.hancom.com/webhwp/devguide/hwpctrl/methods)
- [pyhwpx - 찾기/바꾸기](https://pyhwpx.com)

### Run 기반 편집 액션

| 액션 | 설명 |
|------|------|
| `Delete` | 커서 뒤 글자 삭제 (선택 영역 있으면 선택 삭제) |
| `BackSpace` | 커서 앞 글자 삭제 |
| `Cut` | 선택 영역 잘라내기 |
| `Copy` | 선택 영역 복사 |
| `Paste` | 붙여넣기 |
| `SelectAll` | 전체 선택 |
| `Undo` | 되돌리기 |
| `Redo` | 다시 실행 |

### AllReplace (찾기/바꾸기)

HParameterSet.HFindReplace 사용:

```js
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "찾을문자";
hwp.HParameterSet.HFindReplace.ReplaceString = "바꿀문자";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;  // 대화상자 억제
hwp.HParameterSet.HFindReplace.ReplaceMode = 1;    // 전체 바꾸기
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
```

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 | 실험 결과 | 비고 |
|------|-----------|-----------|------|
| AllReplace 빈 문자열 | 빈 문자열로 치환 (삭제 효과) | **0x80010105 에러** | 보안모듈 관련 에러코드 |
| BackSpace | 커서 앞 글자 삭제 | 문서 끝에서 안 먹힘 | 특수 상황 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)
- 보안모듈 미설치 상태

### 선택 + 삭제/교체

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 4-1 | MoveSelRight x3 → Delete | ✅ | 선택 영역 삭제 정상 |
| 4-2 | MoveSelRight x3 → InsertText | ✅ | 선택 영역 덮어쓰기 정상 |
| 4-3 | 위치 지정 교체 (MoveRight + MoveSelRight + InsertText) | ✅ | 프로그래밍으로 부분 교체 가능 |
| 4-4 | SelectAll → Delete | ✅ | 전체 삭제 |
| 4-5 | Cut | ✅ | 잘라내기 정상 |

### AllReplace (찾기/바꾸기)

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 4-6 | AllReplace (HParameterSet) | ✅ | "BB" → "XX" 정상 |
| 4-7 | AllReplace (CreateAction) | ✅ | "Hello" → "Hi" 정상 |
| 4-8 | AllReplace 한글 | ✅ | "가나다" → "XYZ" 정상 |
| 4-9 | AllReplace 여러 문단 | ✅ | 문단 경계 넘어 전체 치환 |
| 4-10 | AllReplace 공백으로 치환 | ✅ | "---" → " " 정상 |
| 4-11 | AllReplace **빈 문자열** | ❌ | 0x80010105 에러 |

### 전체 교체 패턴

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 4-12 | SelectAll + Delete + SetTextFile | ✅ | 안정적 전체 교체 |

### Undo / Redo

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 4-13 | Undo | ✅ | Delete 후 원래 텍스트 복원 |
| 4-14 | Redo | ✅ | Undo 후 다시 적용 |

---

## 4. 분석

### 핵심 발견 1: 텍스트 수정 — 두 가지 방법 모두 가능

**방법 A: Run 기반 선택 (MoveSelRight 등)**
```js
hwp.Run("MoveDocBegin");
for (var i = 0; i < 3; i++) hwp.Run("MoveRight");     // 위치 이동
for (var i = 0; i < 3; i++) hwp.Run("MoveSelRight");   // 범위 선택
hwp.InsertText("xxx");                                  // 덮어쓰기
```

**방법 B: SelectText + pos 인코딩** (단계 3에서 발견)
```js
var offset = getParaOffset(0);  // 동적 오프셋 감지
hwp.SelectText(0, offset+3, 0, offset+6);  // 4~6번째 글자 선택
hwp.InsertText("xxx");                       // 덮어쓰기
```

SelectText는 이전에 "가짜 선택"으로 판단했으나, **pos 인코딩 문제였을 뿐 정상 작동한다** (단계 3 분석 참고).

### 핵심 발견 2: AllReplace는 강력하지만 빈 문자열 불가

찾기/바꾸기는 **한글, 영문, 여러 문단 모두** 정상 작동.
`IgnoreMessage = 1` 로 대화상자 억제 가능.
두 패턴 (HParameterSet, CreateAction) 모두 작동.

**단, ReplaceString을 빈 문자열(`""`)로 하면 0x80010105 에러.**
삭제 효과가 필요하면 공백(`" "`)으로 치환하거나, Find → 선택 → Delete 패턴 사용.

### 핵심 발견 3: 문서 수정 패턴 정리

| 작업 | 방법 A (Run 기반) | 방법 B (SelectText) |
|------|-------------------|-------------------|
| 특정 위치 텍스트 삭제 | MoveRight x N → MoveSelRight x M → Delete | SelectText(para, offset+start, para, offset+end) → Delete |
| 특정 위치 텍스트 교체 | MoveRight x N → MoveSelRight x M → InsertText | SelectText + InsertText |
| 찾아서 바꾸기 | AllReplace (IgnoreMessage=1) | 동일 |
| 찾아서 삭제 | AllReplace로 공백 치환 | 동일 |
| 전체 교체 | SelectAll → Delete → SetTextFile | 동일 |
| 되돌리기/다시 | Undo / Redo | 동일 |

---

## 5. 최종 결론

### 정상 작동 (✅)

- **Run 기반 선택 + Delete/InsertText** — 선택 영역 삭제/덮어쓰기
- **SelectText + Delete/InsertText** — pos 인코딩 적용 시 정상 작동
- **AllReplace** — 찾기/바꾸기 (한글, 영문, 여러 문단, IgnoreMessage)
- **Cut** — 잘라내기
- **SelectAll + Delete + SetTextFile** — 전체 교체
- **Undo / Redo** — 되돌리기/다시 실행

### 작동하지 않음 (❌)

- **AllReplace 빈 문자열** — 0x80010105 에러. 공백 치환으로 우회
