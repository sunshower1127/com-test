# 한글(HWP) 자동화 스킬 프롬프트

## 두 가지 API 패턴 (둘 다 사용 가능)

**패턴 1: HParameterSet (직접 프로퍼티)**

```js
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = "안녕하세요";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
```

**패턴 2: CreateAction (SetItem 딕셔너리)**

```js
var act = hwp.CreateAction("InsertText");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Text", "안녕하세요");
act.Execute(set);
```

둘 다 정상 작동한다. 상황에 맞게 선택.

---

## ⚠️ 핵심 주의사항 (Gotchas)

### 1. pos 인코딩 — 글자 인덱스가 아님

SetPos, SelectText 등에서 pos 값은 HWP 내부 오프셋이며, **글자 인덱스와 다르다.** 문단마다 오프셋이 다를 수 있으므로, **반드시 동적으로 읽어서 사용할 것.**

```js
// 문단의 시작 오프셋을 동적으로 읽기
var getParaOffset = function(para) {
  hwp.SetPos(0, para, 0);
  hwp.Run("MoveParaBegin");
  return hwp.GetPosBySet().Item("Pos");
};

// 해당 문단의 N번째 글자로 이동
var offset = getParaOffset(0);  // 첫 문단 (보통 16이지만 보장 안 됨)
hwp.SetPos(0, 0, offset + 5);  // 5번째 글자로

// 두번째 문단
var offset1 = getParaOffset(1);  // 보통 0이지만 동적으로 확인
hwp.SetPos(0, 1, offset1 + 5);
```

### 2. SelectText도 pos 인코딩 필요

```js
var offset = getParaOffset(0);
hwp.SelectText(0, offset, 0, offset + 5);  // ✅ 첫 문단 0~4번째 글자
hwp.SelectText(0, 0, 0, 5);                // ❌ 빈 결과 (오프셋 무시)
```

### 3. SetTextFile은 "교체"가 아닌 "삽입"

```js
// ✅ 우회: 먼저 전체 삭제 후 삽입
hwp.Run("SelectAll");
hwp.Run("Delete");
hwp.SetTextFile(newContent, "UNICODE", "");
```

### 4. BreakPara가 자동 맞춤법 교정을 트리거

```js
// ⚠️ "첫번째" → "첫 번째"로 자동 교정됨
hwp.Run("BreakPara");

// ✅ \r\n으로 대체하면 교정 안 됨
set.SetItem("Text", "첫번째\r\n두번째");
```

### 5. 표 셀 배경색 — WinBrushFaceStyle=-1 필수

```js
hwp.HAction.GetDefault("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
var fa = hwp.HParameterSet.HCellBorderFill.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = hwp.RGBColor(255, 0, 0);
fa.WinBrushFaceStyle = -1;  // 반드시 -1! (0이나 1은 빗금 패턴)
hwp.HAction.Execute("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
```

WinBrushFaceStyle 값: -1=단색, 0=가로줄, 1=세로줄, 2~5=기타 빗금 (Windows GDI HBRUSH 규격)

### 6. 이미지는 로컬 파일만 가능

```js
hwp.InsertPicture("C:/path/to/image.jpg", 1, 0, 0, 0, 0, 0, 0);  // ✅
// URL 불가, base64 불가, HTML img 불가
```

### 7. 색상은 RGBColor 함수 사용

```js
hwp.RGBColor(255, 0, 0);  // 빨강
hwp.RGBColor(0, 0, 255);  // 파랑
// 내부적으로 Windows COLORREF(BGR)로 변환되지만 신경 쓸 필요 없음
```

### 8. SaveAs는 보안모듈 필요

Open은 보안 팝업만 허용하면 작동. SaveAs는 보안모듈(DLL) 설치 필요. 없으면 `0x80010105` 에러.

### 9. GetTextFile 포맷별 차이

| 포맷 | 용도 | 비고 |
|------|------|------|
| `"UNICODE"` | 순수 텍스트 | 가볍고 빠름 |
| `"HWPML2X"` | 전체 XML (서식 포함) | 라운드트립 가능, 유일하게 수정 후 재삽입 가능 |
| `"HTML"` | HTML | CP949 인코딩, 라운드트립 불가 |
| `"UTF8"` | — | 미지원 (null 반환) |

---

## 텍스트 추출

```js
var all = hwp.GetTextFile("UNICODE", "");           // 전체 문서
hwp.Run("MoveSelLineEnd");
var sel = hwp.GetTextFile("UNICODE", "saveblock");  // 선택 영역
var p1 = hwp.GetPageText(0);                         // 1페이지 (0-based)
```

---

## 주요 Action ID

| Action ID | 용도 | 주요 프로퍼티 |
|-----------|------|--------------|
| `InsertText` | 텍스트 삽입 | `Text` |
| `CharShape` | 글자 모양 | `Height`(1/10pt), `Bold`(0/1), `TextColor` |
| `ParagraphShape` | 문단 모양 | `Alignment`(0=양쪽,1=왼쪽,2=오른쪽,3=가운데) |
| `TableCreate` | 표 생성 | `Rows`, `Cols`, `WidthType`(0=단에맞춤) |
| `AllReplace` | 찾기/바꾸기 | `FindString`, `ReplaceString`, `IgnoreMessage`(1) |
| `CellBorderFill` | 셀 서식 | `FillAttr` 하위 |

---

## Run 명령

| 명령 | 동작 |
|------|------|
| `MoveDocBegin` / `MoveDocEnd` | 문서 처음/끝 |
| `SelectAll` / `Delete` / `Cancel` | 전체 선택/삭제/선택 해제 |
| `MoveRight` / `MoveLeft` | 한 글자 이동 |
| `MoveSelRight` / `MoveSelLineEnd` / `MoveSelDocEnd` | 선택 확장 |
| `MoveLineBegin` / `MoveLineEnd` | 줄 처음/끝 |
| `MoveParaBegin` / `MoveNextParaBegin` | 문단 처음/다음 문단 |
| `MoveDown` / `MoveUp` | 아래/위로 (표 진입 가능) |
| `TableRightCell` / `TableLeftCell` | 표 셀 좌우 이동 |
| `TableUpperCell` / `TableLowerCell` | 표 셀 상하 이동 |
| `TableCellBlock` | 현재 셀부터 블록 선택 시작 |
| `FileNew` | 새 문서 |

---

## 기존 문서 수정 전략

### 전략 1: AllReplace (텍스트 치환)

가장 간단하고 안전. 사용자 수정도 보존.

```js
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "홍길동";
hwp.HParameterSet.HFindReplace.ReplaceString = "김철수";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
```

### 전략 2: HWPML2X 라운드트립 (표 셀 수정 등)

복잡한 표의 특정 셀 수정에 사용. 서식 100% 보존.

```
1. GetTextFile("HWPML2X", "") → 전체 XML
2. XML 파싱 → 구조맵 (LLM에게 전달)
3. LLM이 좌표+값 지정 → 코드가 XML에서 해당 CELL 수정
4. SelectAll → Delete → SetTextFile("HWPML2X") → 서식 보존 라운드트립
```

**주의:** 수정 직전에 XML을 다시 가져올 것 (사용자가 중간에 수정했을 수 있음)

### 전략 3: 이미지 삽입 (플레이스홀더 패턴)

```
1. HWPML2X에 플레이스홀더 텍스트 삽입 ("{{PHOTO}}")
2. 라운드트립으로 반영
3. Find("{{PHOTO}}")로 커서 이동
4. Delete로 플레이스홀더 삭제
5. InsertPicture로 이미지 삽입
```

---

## 표 접근

### 표 진입 (MoveDown)

MoveDown으로 어떤 표든 진입 가능. 단, **도착 셀은 예측 불가능.**

```js
hwp.Run("MoveDocBegin");
for (var i = 0; i < 20; i++) {
  hwp.Run("MoveDown");
  var list = hwp.GetPosBySet().Item("List");
  if (list > 0) break;  // List > 0 = 표 안
}
// TableUpperCell/LeftCell로 원하는 셀로 이동
```

### 셀 내 수정

```js
// 셀 안에서 전체 선택 → 삭제 → 새 텍스트
hwp.Run("SelectAll");   // 표 안에서는 현재 셀 텍스트만 선택됨
hwp.Run("Delete");
// InsertText로 새 내용 입력
```

### List ID 직접 점프

각 셀은 고유한 List ID를 가진다. 알고 있으면 직접 점프 가능:

```js
hwp.SetPos(listId, 0, 0);  // 해당 셀로 즉시 이동
```

단, List ID는 런타임에만 존재하고 문서마다 다름. HWPML2X/HTML에는 없음.

---

## 커서 위치 읽기

```js
var ps = hwp.GetPosBySet();
var list = ps.Item("List");  // 0=본문, >0=표 셀 등
var para = ps.Item("Para");
var pos = ps.Item("Pos");
```

위치 저장/복원:

```js
var saved = hwp.GetPosBySet();
// ... 다른 작업 ...
hwp.SetPosBySet(saved);  // 원래 위치로 복귀
```

---

## xml 객체 — HWPML2X 편집 API

문서의 HWPML2X를 로드하고, 경로 지정으로 부분 수정한 뒤 한번에 반영하는 도구.

### 기본 사용법

```js
xml.load()                           // 현재 문서의 HWPML2X 로드
result = xml.structure()             // 구조맵 확인 (표/본문 목록)

xml.set("t0.r1.c3", "홍길동")        // 표0, 행1, 열3 전체 텍스트 교체
xml.set("t0.r1.c6", "1990.01.15")   // 표0, 행1, 열6
xml.set("p2", "수정된 본문")          // 본문 문단2 전체 교체

result = xml.commit()                // 수정사항 문서에 반영
```

### 경로 문법

| 경로 | 의미 |
|------|------|
| `t0.r1.c3` | 표0의 행1, 열3 (ColAddr 기준) — 전체 텍스트 교체 |
| `t0.r1.c3.p0` | 표0의 행1, 열3 — 셀 내 첫번째 문단만 |
| `t0.r1.c3.p1` | 표0의 행1, 열3 — 셀 내 두번째 문단만 |
| `p0` | 본문 첫번째 문단 (표 밖) |
| `p3` | 본문 네번째 문단 |

- `t` = table, `r` = row, `c` = col (ColAddr), `p` = paragraph
- col 번호는 배열 인덱스가 아닌 **ColAddr** (병합 셀 때문에 불연속일 수 있음)
- 여러 줄 텍스트: `xml.set("t0.r1.c3", "줄1\r\n줄2\r\n줄3")`

### 구조맵 예시

```
=== Body Paragraphs ===
  p0: 이  력  서
  p1: 지원분야:                     개발직

=== Tables: 2 ===

--- t0 (24 rows) ---
  r0: | c0:(empty) [2x4] | c2: 성       명 | c3:(empty) [2x1] | c5: 생 년 월 일 | c6:(empty) [2x1] |
  r1: | c2: 연 락 처 | c3:(empty) [2x1] | c5:비상연락망 | c6:(empty) [2x1] |
  ...

--- t1 (4 rows) ---
  r0: | c0:성장과정 | c1:(empty) |
  ...
```

### 전체 흐름 예시

```js
// 이력서 열고 정보 채우기
hwp.Open("C:/path/이력서.hwp", "HWP", "")

xml.load()
result = xml.structure()  // 구조 확인

// 1페이지: 이력서 표
xml.set("t0.r0.c3", "홍길동")
xml.set("t0.r0.c6", "1990.01.15")
xml.set("t0.r1.c3", "010-1234-5678")
xml.set("t0.r1.c6", "010-9876-5432")
xml.set("t0.r2.c3", "hong@example.com")
xml.set("t0.r3.c3", "서울시 강남구 테헤란로 123")

// 2페이지: 자기소개서 표
xml.set("t1.r0.c1", "어릴 때부터 컴퓨터에 관심이 많았습니다...")
xml.set("t1.r1.c1", "꼼꼼하고 책임감이 강합니다...")

result = xml.commit()  // 한번에 반영

// 이미지 삽입 (플레이스홀더 패턴)
// xml.set으로 사진 셀에 "{{PHOTO}}" 넣고 commit 후:
hwp.HAction.GetDefault("FindReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "{{PHOTO}}";
hwp.HParameterSet.HFindReplace.FindType = 1;
hwp.HAction.Execute("FindReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.Run("Delete");
hwp.InsertPicture("C:/path/photo.jpg", 1, 0, 0, 0, 0, 0, 0);
```

### 주의사항

- **commit 전에 xml.load()** 필수 — 사용자가 중간에 문서를 수정했을 수 있음
- **commit은 전체 문서 교체** — HWPML2X 전체를 SelectAll+Delete+SetTextFile
- **col 번호는 ColAddr** — 병합 셀이 있으면 c0 다음이 c2일 수 있음 (structure로 확인)
- **xml.set은 큐에 쌓임** — commit() 호출 전까지 문서에 반영 안 됨
