# HWP COM 팁

전역 객체: `hwp` (한글 COM Proxy), `xml` (HWPML2X 편집)

결과를 반환하려면 반드시 `result = ...` 형태로 할당할 것. 할당하지 않으면 출력에 표시되지 않음.

## 기본 흐름

**hwp 블록과 xml 블록은 반드시 분리한다.** 같은 블록에 `hwp.*`과 `xml.*`을 섞지 않는다.

```
[hwp 블록] 문서 열기 (파일 경로가 주어진 경우)
hwp.Open("C:/path/to/file.hwp", "HWP", "");

[xml 블록] 1단계: 문서 파악
result = xml.outline()
```

**1단계 결과를 보고 판단:**
- outline 결과가 짧으면 (수십 줄 이내) → `xml.get()`으로 전체를 한번에 읽기
- outline 결과가 길면 → 사용자에게 어느 부분을 볼지 질문
  - 예: "문서가 31페이지입니다. 어떤 페이지/부분부터 확인할까요?"

```
[xml 블록] 2단계: 상세 확인
result = xml.get()           // 전체 (작은 문서)
result = xml.get("t0")       // 특정 표
result = xml.get("p0~p10")   // 특정 범위

[hwp 블록] 3단계: 수정 작업
...
```

**주의사항:**
- **hwp.Open과 xml.* 호출을 같은 블록에 넣지 말 것** — Open 직후 xml 호출하면 문서 로드가 완료되지 않아 빈 결과가 나옴
- xml 블록은 시작 시 자동 load, 종료 시 자동 commit
- 파일 경로가 주어지면 반드시 `hwp.Open()`부터 시작. 파일 경로가 없으면 이미 열린 문서가 있다고 가정하고 바로 `xml.outline()`부터 시작
- hwp.Open 없이도 xml.outline()은 동작함 (이미 열려있는 문서 대상)

## ⚠️ 핵심 주의사항 (Gotchas)

### 1. pos 인코딩 — 글자 인덱스가 아님

SetPos, SelectText 등에서 pos 값은 HWP 내부 오프셋이며, **글자 인덱스와 다르다.** 문단마다 오프셋이 다를 수 있으므로, **반드시 동적으로 읽어서 사용할 것.**

```js
// 문단의 시작 오프셋을 동적으로 읽기
var getParaOffset = function (para) {
  hwp.SetPos(0, para, 0);
  hwp.Run("MoveParaBegin");
  return hwp.GetPosBySet().Item("Pos");
};

// 해당 문단의 N번째 글자로 이동
var offset = getParaOffset(0); // 첫 문단 (보통 16이지만 보장 안 됨)
hwp.SetPos(0, 0, offset + 5); // 5번째 글자로

// 두번째 문단
var offset1 = getParaOffset(1); // 보통 0이지만 동적으로 확인
hwp.SetPos(0, 1, offset1 + 5);
```

### 2. SelectText도 pos 인코딩 필요

```js
var offset = getParaOffset(0);
hwp.SelectText(0, offset, 0, offset + 5); // ✅ 첫 문단 0~4번째 글자
hwp.SelectText(0, 0, 0, 5); // ❌ 빈 결과 (오프셋 무시)
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
hwp.HAction.GetDefault(
  "CellBorderFill",
  hwp.HParameterSet.HCellBorderFill.HSet
);
var fa = hwp.HParameterSet.HCellBorderFill.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = hwp.RGBColor(255, 0, 0);
fa.WinBrushFaceStyle = -1; // 반드시 -1! (0이나 1은 빗금 패턴)
hwp.HAction.Execute("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
```

WinBrushFaceStyle 값: -1=단색, 0=가로줄, 1=세로줄, 2~5=기타 빗금 (Windows GDI HBRUSH 규격)

### 6. 이미지는 로컬 파일만 가능

```js
hwp.InsertPicture("C:/path/to/image.jpg", 1, 0, 0, 0, 0, 0, 0); // ✅
// URL 불가, base64 불가, HTML img 불가
```

### 7. 색상은 RGBColor 함수 사용

```js
hwp.RGBColor(255, 0, 0); // 빨강
hwp.RGBColor(0, 0, 255); // 파랑
// 내부적으로 Windows COLORREF(BGR)로 변환되지만 신경 쓸 필요 없음
```

### 8. SaveAs는 보안모듈 필요

Open은 보안 팝업만 허용하면 작동. SaveAs는 보안모듈(DLL) 설치 필요. 없으면 `0x80010105` 에러.

### 9. GetTextFile 포맷별 차이

| 포맷        | 용도                 | 비고                                          |
| ----------- | -------------------- | --------------------------------------------- |
| `"UNICODE"` | 순수 텍스트          | 가볍고 빠름                                   |
| `"HWPML2X"` | 전체 XML (서식 포함) | 라운드트립 가능, 유일하게 수정 후 재삽입 가능 |
| `"HTML"`    | HTML                 | CP949 인코딩, 라운드트립 불가                 |
| `"UTF8"`    | —                    | 미지원 (null 반환)                            |

---

## 텍스트 추출

```js
var all = hwp.GetTextFile("UNICODE", ""); // 전체 문서
hwp.Run("MoveSelLineEnd");
var sel = hwp.GetTextFile("UNICODE", "saveblock"); // 선택 영역
var p1 = hwp.GetPageText(0); // 1페이지 (0-based)
```

---

## 찾기 — RepeatFind

⚠️ **찾기는 `RepeatFind`만 사용할 것.** `FindReplace`, `Find`, `FindDlg` 등 다른 찾기 API는 전부 작동하지 않음 (HWP 자체 제한).

```js
// 기본 패턴
hwp.Run("MoveDocBegin");
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "찾을텍스트";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;  // 필수: 팝업 억제
var found = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
// found=true → 텍스트가 선택된 상태. 바로 Delete, InsertText 등 가능
```

- 본문, 표 안 모두 찾기 가능
- 호출할 때마다 다음 매칭으로 이동
- **순환 주의**: 문서 끝 도달 시 처음으로 돌아가며 true 반환. 위치 추적으로 무한루프 방지 필요

### 플레이스홀더 → 이미지 삽입 패턴

```js
// 1. xml.set으로 셀에 플레이스홀더
xml.load();
xml.set("t0.r1.c0", "{{PHOTO}}");
xml.commit(); // 반드시 commit 후 RepeatFind

// 2. RepeatFind로 찾아서 이미지 교체
hwp.Run("MoveDocBegin");
hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "{{PHOTO}}";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
hwp.Run("Delete");
var ctrl = hwp.InsertPicture("C:/path/photo.jpg", 1, 0, 0, 0, 0, 0, 0);
// 크기 조절 (HWPUNIT: 1mm ≈ 283.46)
if (ctrl) {
    var props = ctrl.Properties;
    props.SetItem("Width", 8041);  // ~28mm
    props.SetItem("Height", 10440); // ~37mm
    ctrl.Properties = props;
}
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
  if (list > 0) break; // List > 0 = 표 안
}
// TableUpperCell/LeftCell로 원하는 셀로 이동
```

### 셀 내 수정

```js
// 셀 안에서 전체 선택 → 삭제 → 새 텍스트
hwp.Run("SelectAll"); // 표 안에서는 현재 셀 텍스트만 선택됨
hwp.Run("Delete");
// InsertText로 새 내용 입력
```

### List ID 직접 점프

각 셀은 고유한 List ID를 가진다. 알고 있으면 직접 점프 가능:

```js
hwp.SetPos(listId, 0, 0); // 해당 셀로 즉시 이동
```

단, List ID는 런타임에만 존재하고 문서마다 다름. HWPML2X/HTML에는 없음.

---

## 커서 위치 읽기

```js
var ps = hwp.GetPosBySet();
var list = ps.Item("List"); // 0=본문, >0=표 셀 등
var para = ps.Item("Para");
var pos = ps.Item("Pos");
```

위치 저장/복원:

```js
var saved = hwp.GetPosBySet();
// ... 다른 작업 ...
hwp.SetPosBySet(saved); // 원래 위치로 복귀
```

---

## xml 객체 — HWPML2X 편집 API

문서의 HWPML2X를 로드하고, 구조 파악 및 수정을 하는 도구. 두 가지 모드가 있다:

1. **set/commit**: 텍스트만 빠르게 교체 (대량 입력에 적합)
2. **raw/rawSet**: 원본 XML을 직접 보고 수정 (스타일/이미지 등 세밀한 제어)

### 모드 1: set/commit — 텍스트 빠른 교체

```js
xml.load();
result = xml.structure(); // 구조맵 확인

xml.set("t0.r1.c3", "홍길동");
xml.set("t0.r1.c6", "1990.01.15");
xml.set("p2", "수정된 본문");
xml.commit(); // 반드시 호출
```

### 모드 2: raw/rawSet — XML 직접 수정

```js
xml.load();
result = xml.raw("t0.r1.c3"); // 해당 셀의 원본 XML 반환
// → <CELL ColAddr="3" ...><TEXT CharShape="2"><CHAR>홍길동</CHAR></TEXT></CELL>

// XML을 직접 수정해서 교체 (스타일, 이미지 크기, 서식 등 자유롭게)
xml.rawSet("t0.r1.c3", '<CELL ColAddr="3" ...><TEXT CharShape="2"><CHAR>김철수</CHAR></TEXT></CELL>');
// → 즉시 문서에 반영 (commit 불필요)
```

**rawSet은 즉시 반영**됨. commit이 필요 없음. XML 태그 균형이 맞지 않으면 경고 표시.

### 언제 어떤 모드?

| 상황 | 모드 |
|------|------|
| 빈 칸에 텍스트 대량 입력 | `set/commit` |
| 이미지 크기, 셀 서식 수정 | `raw/rawSet` |
| 스타일 속성 변경 | `raw/rawSet` |
| 기존 텍스트 부분 수정 | `raw/rawSet` |

### 경로 문법

| 경로          | 의미                                             |
| ------------- | ------------------------------------------------ |
| `t0.r1.c3`    | 표0의 행1, 열3 (ColAddr 기준) |
| `t0.r1.c3.p0` | 표0의 행1, 열3 — 셀 내 첫번째 문단만 (set 전용) |
| `p0`          | 본문 첫번째 문단 (표 밖)                         |
| `p3`          | 본문 네번째 문단                                 |

- `t` = table, `r` = row, `c` = col (ColAddr), `p` = paragraph
- col 번호는 배열 인덱스가 아닌 **ColAddr** (병합 셀 때문에 불연속일 수 있음)

### 구조맵 (structure)

마커 형식으로 문서 순서 유지. 문단과 표가 섞여서 표시됨:

```
---p0---
이  력  서
---!p0---
---t0 (24 rows, 8 cols)---

---t0.r1.c2---
성       명
---!t0.r1.c2---
---t0.r1.c3 [2x1]---

---!t0.r1.c3---
...
---!t0---
---p2---
위 내용은 사실과 틀림없음.
---!p2---
```

### 주의사항

- **set 전에 xml.load()** 필수 — 사용자가 중간에 문서를 수정했을 수 있음
- **set 후에 xml.commit()** 필수 — commit 후에 다른 작업(RepeatFind 등)을 해야 정상 동작
- **rawSet은 commit 불필요** — 즉시 반영
- **raw 수정 시 load는 직전에** — 사용자가 문서를 수정했을 수 있으므로 raw 전에 load
- **내부적으로 전체 문서 교체** — HWPML2X 전체를 SelectAll+Delete+SetTextFile
- **col 번호는 ColAddr** — 병합 셀이 있으면 c0 다음이 c2일 수 있음 (structure로 확인)
- **이미지가 있는 셀에 set()** — PICTURE 태그가 제거됨. 이미지 조작은 raw/rawSet 사용
