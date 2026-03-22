# HWP COM Cheat Sheet (레퍼런스)

hwp-tips.md의 기본 규칙을 숙지한 상태에서 참고용으로 사용.

## CharShape 속성

| 속성           | 값                            |
| -------------- | ----------------------------- |
| Height         | 1/10pt (1000=10pt, 2400=24pt) |
| Bold           | 0/1                           |
| Italic         | 0/1                           |
| TextColor      | `hwp.RGBColor(r,g,b)`         |
| FaceNameHangul | "맑은 고딕"                   |
| FaceNameLatin  | "Arial"                       |
| UnderlineType  | 0=없음, 1=밑줄                |
| StrikeOutType  | 0=없음                        |
| SuperScript    | 0/1                           |
| SubScript      | 0/1                           |
| Spacing        | -50~50 (자간)                 |

## ParagraphShape 속성

| 속성              | 값                                 |
| ----------------- | ---------------------------------- |
| Alignment         | 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데 |
| LineSpacing       | % (160=160%)                       |
| LineSpacingType   | 0=%, 1=고정, 2=여백, 3=최소        |
| ParaMarginLeft    | hu단위 (283.46hu = 1mm)            |
| ParaMarginRight   | hu                                 |
| Indent            | hu                                 |
| ParaSpacingBefore | hu                                 |
| ParaSpacingAfter  | hu                                 |

## AllReplace 속성

| 속성          | 값            |
| ------------- | ------------- |
| FindString    | "찾을것"      |
| ReplaceString | "바꿀것"      |
| IgnoreMessage | 1 (팝업 억제) |
| MatchCase     | 0/1           |
| WholeWordOnly | 0/1           |
| FindRegExp    | 0/1           |

## TableCreate 속성

| 속성      | 값                           |
| --------- | ---------------------------- |
| Rows      | 행 수                        |
| Cols      | 열 수                        |
| WidthType | 0=단맞춤, 1=문단맞춤, 2=지정 |

표 속성 하위: `var tp = set.Item("TableProperties"); tp.SetItem("TreatAsChar", 1); set.SetItem("TableProperties", tp)`

## PageSetup 속성

하위 객체: `var pd = set.Item("PageDef")`

| 속성 | 값 |
|------|-----|
| LeftMargin | hu (4252≈15mm, 5669≈20mm) |
| RightMargin | hu |
| TopMargin | hu |
| BottomMargin | hu |
| Landscape | 0=세로, 1=가로 |

`pd.SetItem("LeftMargin", 5669); set.SetItem("PageDef", pd)`

## 기타 Action ID

| ID               | 용도           |
| ---------------- | -------------- |
| BorderFill       | 테두리/배경    |
| CellBorderFill   | 셀 테두리/배경 |
| HeaderFooter     | 머리글/꼬리글  |
| FootNote         | 각주           |
| HyperlinkPrivate | 하이퍼링크     |
| InsertFile       | 파일 삽입      |

## 직접 메서드

```js
hwp.Open(path, "HWP", "");
hwp.SaveAs(path, "HWP", ""); // 3인자 필수
hwp.Clear(1); // 저장 안 하고 닫기
hwp.GetTextFile("UNICODE", ""); // 전체 추출
hwp.GetTextFile("UNICODE", "saveblock"); // 선택 영역
hwp.InsertPicture(path, true, sizeoption, false, false, 0, width_mm, height_mm);
hwp.MovePos(2); // 2=문서시작, 3=문서끝
hwp.RGBColor(r, g, b);
hwp.MiliToHwpUnit(20.0); // mm→hu
```

포맷: `"HWP"`, `"HWPX"`, `"PDF"`, `"UNICODE"`, `"HTML"`, `"OOXML"`

## 필드 (누름틀)

```js
hwp.PutFieldText("이름", "홍길동");
var text = hwp.GetFieldText("이름{{0}}");
hwp.MoveToField("이름", true, true, false);
```

## 단위 변환

| 변환        | 공식                    |
| ----------- | ----------------------- |
| mm → hu     | × 283.46                |
| hu → mm     | ÷ 283.46                |
| pt → Height | × 100 (10pt=1000)       |
