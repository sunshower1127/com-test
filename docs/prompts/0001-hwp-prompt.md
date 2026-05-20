# HWP COM 팁

전역 객체: `hwp` (한글 COM Proxy), `xml` (HWPML2X 편집)

**hwp 블록과 xml 블록은 반드시 분리한다.** 같은 블록에 `hwp.*`과 `xml.*`을 섞지 않는다.

결과를 반환하려면 반드시 `result = ...` 형태로 할당할 것.

## 기본 흐름

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

### 3. 텍스트 추가/교체 방법

**hwp (COM 직접):**

```js
// 커서 위치에 텍스트 삽입
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = "새 텍스트";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);

// 여러 줄 삽입 — 반드시 \r\n으로 개행
hwp.HParameterSet.HInsertText.Text = "첫째 줄\r\n둘째 줄\r\n셋째 줄";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);

// 선택 영역 교체: 먼저 선택 → InsertText (선택 영역을 덮어씀)
hwp.Run("MoveDocBegin");
hwp.Run("MoveRight"); // 원하는 위치로
hwp.Run("Select"); // 선택 시작
hwp.Run("MoveRight"); // 선택 범위
hwp.HParameterSet.HInsertText.Text = "교체 텍스트";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
```

**xml (구조적 편집):**

```js
// 특정 경로의 텍스트 교체 (서식 유지)
xml.setText("p0", "새 텍스트"); // 본문 문단
xml.setText("t0.r1.c0", "셀 텍스트"); // 표 셀
xml.setText("t0.r1.c0.p0", "첫째 문단"); // 셀 내 특정 문단

// 원본 XML 교체 (서식 포함)
xml.rawSet("p0", "<P ...>새 XML</P>");

// 요소 삽입
xml.insert("p0", "<P ...>뒤에 추가</P>"); // p0 뒤에 삽입
xml.insertBefore("p0", "<P ...>앞에 추가</P>"); // p0 앞에 삽입
```

> **⚠️ SetTextFile 금지** — hwp 코드에서 `SetTextFile`을 사용하지 마세요. 이 함수는 "교체"가 아닌 "삽입"으로 동작하여 문서가 이중으로 됩니다. 텍스트 삽입은 `InsertText`, 구조적 편집은 `xml.setText`를 사용하세요. (`SetTextFile`은 xml 내부의 HWPML2X 로드에서만 사용됩니다.)

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

### 8. SaveAs는 보안모듈 필요

Open은 보안 팝업만 허용하면 작동(떠있는 동안은 블록킹). SaveAs는 보안모듈(DLL) 설치 필요. 없으면 `0x80010105` 에러.

### 9. GetTextFile은 UNICODE 전용

```js
var text = hwp.GetTextFile("UNICODE", "");           // 문서 전체
var selected = hwp.GetTextFile("UNICODE", "saveblock"); // 선택 영역만
```

> **사용 금지 포맷:**
> - `"UTF8"` — null 반환 (미지원)
> - `"HTML"` — 인코딩 깨짐
> - `"HWPML2X"`, `"HWP"` — hwp에서 직접 사용하지 마세요. 구조적 읽기는 `xml.get()`, `xml.outline()`을 사용하세요.

**텍스트 읽기 방법 비교:**

| 목적 | hwp | xml |
|------|-----|-----|
| 문서 전체 텍스트 | `hwp.GetTextFile("UNICODE", "")` | — |
| 선택 영역 텍스트 | `hwp.GetTextFile("UNICODE", "saveblock")` | — |
| 특정 문단 | — | `xml.get("p0")` |
| 표 셀 내용 | — | `xml.get("t0.r1.c0")` |
| 문서 구조 파악 | — | `xml.outline()` |
| 스타일 목록 | — | `xml.styles()` |
| 특정 스타일 상세 | — | `xml.styles("CharShape", 3)` |

### 10. 스타일 조회/추가 (xml)

```js
// 스타일 목록 조회
result = xml.styles()                    // CharShape + ParaShape 전체
result = xml.styles("CharShape")         // CharShape만
result = xml.styles("CharShape", 3)      // CharShape[3] 상세 속성

// 새 스타일 추가 (ID 반환)
var id = xml.addStyle("CharShape", { Height: 1200, Bold: 1 });
// → id를 rawSet 등에서 CharShapeIDRef로 사용
```
