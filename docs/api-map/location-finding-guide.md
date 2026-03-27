# 문서 내 위치 특정 가이드

사용자 요청이 모호할 때("그 표 수정해줘", "3페이지 쪽 좀 고쳐줘") **어디를 수정할지** 특정하는 방법.
아래 순서대로 시도하되, 상황에 맞는 방법을 선택한다.

---

## 방법 1: 사용자에게 직접 물어보기

가장 확실하고 빠른 방법. 모호한 요청에는 항상 먼저 시도.

```
사용자: "그 표부분 수정해줘"
→ "문서에 표가 5개 있어요. 어떤 건가요?"
  - 표0: 목차 (3행, 1페이지)
  - 표1: 실험 결과 (12행, 3페이지)
  - 표2: 참고문헌 (20행, 8페이지)
사용자: "표1"
```

**언제:** 항상 가능. 표/섹션이 여러 개일 때 특히 유용.

---

## 방법 2: 구조 목차 자동 생성

HWPML2X에서 문서 전체 구조를 파악해서 목차를 만든다.

```
[Phase 1] GetTextFile("HWPML2X") → 전체 XML
[Phase 2] 파서가 구조 요약 생성:

문서 구조:
  본문: 15문단
  표0: 이력서 (24행 8열) — 1페이지
  표1: 학력사항 (5행 4열) — 1페이지
  표2: 자기소개서 (4행 2열) — 2페이지
  이미지: 2개
```

**언제:** 문서를 처음 열 때. 문서 전체를 한눈에 파악해야 할 때.
**비용:** GetTextFile 1회 호출 + XML 파싱 (1초 이내)

---

## 방법 3: 텍스트 검색 (Find/AllReplace)

문서 내에서 특정 키워드를 검색해서 위치를 찾는다.

```js
// "실험 결과"라는 텍스트 근처를 찾고 싶을 때
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "실험 결과";
hwp.HParameterSet.HFindReplace.FindType = 1;  // 전방 검색
hwp.HAction.Execute("FindDlg", hwp.HParameterSet.HFindReplace.HSet);
// → 커서가 "실험 결과" 위치로 이동
// → GetPosBySet()으로 현재 위치 확인 가능
```

**언제:** 사용자가 키워드를 언급했을 때 ("실험 결과 부분 고쳐줘")
**장점:** 빠르고 정확. XML 파싱 불필요.
**단점:** 같은 텍스트가 여러 곳에 있으면 어느 건지 모름.

---

## 방법 4: 사용자 드래그 감지

사용자가 직접 HWP에서 영역을 선택(드래그)하면, 그 위치를 읽는다.

```js
// 사용자가 드래그한 후:
var pos = hwp.GetPosBySet();
var list = pos.Item("List");   // 어떤 리스트(셀)에 있는지
var para = pos.Item("Para");   // 몇 번째 문단
var p = pos.Item("Pos");       // 문단 내 위치

// 선택 영역 텍스트도 읽기
var text = hwp.GetTextFile("UNICODE", "saveblock");
```

**언제:** "내가 드래그할 테니 거기 수정해줘" 같은 상호작용 시나리오.
**장점:** 가장 자연스러운 UX. 복잡한 문서에서도 100% 정확.
**단점:** 사용자 개입 필요. 자동화 불가.

---

## 방법 5: 페이지 번호 기반

사용자가 페이지를 지정하면, 해당 페이지의 내용만 추출.

```
사용자: "3페이지 표 고쳐줘"
→ HWPML2X에서 SECTION/페이지 경계 파악
→ 3페이지에 포함된 표만 구조맵 생성
→ LLM에게 해당 표만 전달
```

**GetPageText(pageIndex, "UNICODE") 사용 가능:**
```js
hwp.PageCount;                            // 전체 페이지 수
hwp.GetPageText(2, "UNICODE");            // 3페이지 텍스트 (0-based)
```
- 포맷 파라미터는 무시됨 — 항상 순수 텍스트만 반환
- HWPML2X로 페이지 단위 추출은 불가. 텍스트로 위치 특정 후 HWPML2X 전체에서 해당 노드를 찾는 방식으로 조합

**SelectText로 범위 지정 후 HWPML2X 추출 가능:**
```js
// pos 오프셋 필수 (항목 1 참조)
var offset = getParaOffset(0);
hwp.SelectText(0, offset, 0, offset + 50);
hwp.GetTextFile("HWPML2X", "saveblock");  // 선택 영역만 XML로
```

---

## 방법 6: 이전 컨텍스트 활용

같은 세션에서 이전에 작업한 위치를 기억해서 재활용.

```
사용자: "아까 그 표 다시 수정해줘"
→ 이전에 사용한 table index, row, col 정보 재사용
→ 또는 이전에 저장한 List ID로 SetPos 직접 점프
```

**언제:** 반복 수정 시. "아까 그거" 류의 요청.
**구현:** 세션 내 상태 저장 (table index, 구조맵 캐시, List ID 맵)

---

## 추천 플로우

```
사용자 요청 도착
    ↓
명확한가? ──YES──→ 바로 실행 (방법 3, 5, 6)
    │
    NO
    ↓
구조 목차 생성 (방법 2)
    ↓
목차로 특정 가능? ──YES──→ 해당 영역만 상세 파악 → 실행
    │
    NO
    ↓
사용자에게 질문 (방법 1)
또는 드래그 요청 (방법 4)
```

---

## 위치 특정 후: 수정 실행

위치를 찾았으면 수정 타입에 따라 최적 방법 선택:

| 수정 타입 | 최적 방법 |
|-----------|----------|
| 표 셀 값 변경 | HWPML2X 라운드트립 |
| 본문 텍스트 교체 | AllReplace |
| 본문 텍스트 추가/삭제 | HWPML2X 라운드트립 |
| 이미지 삽입 | 플레이스홀더 + Find + InsertPicture |
| 서식 변경 | HWPML2X 라운드트립 또는 COM 직접 |

자세한 수정 실행 방법은 [usecase-template-fill.md](usecase-template-fill.md) 참조.
