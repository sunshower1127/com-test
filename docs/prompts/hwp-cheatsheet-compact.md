# HWP COM Cheat Sheet (압축)

## 기본 패턴 — 반드시 CreateAction 사용

```js
var act = hwp.CreateAction("ActionID");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Property", value);
act.Execute(set);
```

## Action ID → 주요 속성

### InsertText

`set.SetItem("Text", "내용\r\n줄바꿈")`

### CharShape (글자 모양)

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

### ParagraphShape (문단 모양)

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

### AllReplace (찾기/바꾸기)

| 속성          | 값            |
| ------------- | ------------- |
| FindString    | "찾을것"      |
| ReplaceString | "바꿀것"      |
| IgnoreMessage | 1 (팝업 억제) |
| MatchCase     | 0/1           |
| WholeWordOnly | 0/1           |
| FindRegExp    | 0/1           |

### TableCreate (표 생성)

| 속성      | 값                           |
| --------- | ---------------------------- |
| Rows      | 행 수                        |
| Cols      | 열 수                        |
| WidthType | 0=단맞춤, 1=문단맞춤, 2=지정 |

표 속성 하위: `var tp = set.Item("TableProperties"); tp.SetItem("TreatAsChar", 1); set.SetItem("TableProperties", tp)`

### PageSetup (페이지 설정)

하위 객체: `var pd = set.Item("PageDef")`
| 속성 | 값 |
|------|-----|
| LeftMargin | hu (4252≈15mm, 5669≈20mm) |
| RightMargin | hu |
| TopMargin | hu |
| BottomMargin | hu |
| Landscape | 0=세로, 1=가로 |

`pd.SetItem("LeftMargin", 5669); set.SetItem("PageDef", pd)`

### 기타 Action ID

| ID               | 용도           | ParameterSet      |
| ---------------- | -------------- | ----------------- |
| BorderFill       | 테두리/배경    | HBorderFill       |
| CellBorderFill   | 셀 테두리/배경 | HCellBorderFill   |
| ColDef           | 단 설정        | HColDef           |
| HeaderFooter     | 머리글/꼬리글  | HHeaderFooter     |
| FootNote         | 각주           | HFootEndNote      |
| HyperlinkPrivate | 하이퍼링크     | HHyperlinkPrivate |
| Print            | 인쇄           | HPrint            |
| InsertFile       | 파일 삽입      | HInsertFile       |
| EquationCreate   | 수식           | HEquation         |

## Run 명령 (파라미터 없음)

### 이동

```
MoveDocBegin, MoveDocEnd, MoveLineBegin, MoveLineEnd
MoveParaBegin, MoveParaEnd, MoveNextPara, MovePrevPara
MovePageDown, MovePageUp, MoveWordRight, MoveWordLeft
```

### 선택 (Move + Sel)

```
MoveSelDocBegin, MoveSelDocEnd, MoveSelLineBegin, MoveSelLineEnd
MoveSelNextPara, MoveSelPrevPara, MoveSelPageDown, MoveSelPageUp
```

### 편집

```
SelectAll, Copy, Cut, Paste, Delete, Undo, Redo
```

### 줄바꿈/구역

```
BreakPara, BreakPage, BreakSection, BreakColumn
```

### 표 내 이동

```
TableRightCell, TableLeftCell, TableUpperCell, TableLowerCell
TableAppendRow, TableDeleteRow, TableSplitCell, TableMergeCell
```

### 파일

```
FileNew, FileSave, FileSaveAs, FileClose
```

## 직접 메서드 (CreateAction 아님)

### 파일 I/O

```js
hwp.Open(path, "HWP", "");
hwp.SaveAs(path, "HWP", ""); // 3인자 필수
hwp.Clear(1); // 저장 안 하고 닫기
hwp.Quit();
```

포맷: `"HWP"`, `"HWPX"`, `"PDF"`, `"UNICODE"`, `"HTML"`, `"OOXML"`

### 텍스트 추출

```js
result = hwp.GetTextFile("UNICODE", ""); // 전체
result = hwp.GetTextFile("UNICODE", "saveblock"); // 선택 영역
```

### 필드 (누름틀)

```js
hwp.CreateField("안내문", "", "필드이름");
var fields = hwp.GetFieldList(1, 0); // \x02 구분
hwp.PutFieldText("이름", "홍길동"); // 전체 매칭
hwp.PutFieldText("이름{{0}}", "홍길동"); // 특정 인스턴스
var text = hwp.GetFieldText("이름{{0}}");
hwp.MoveToField("이름", true, true, false); // text, start, select
```

### 이미지 삽입

```js
hwp.InsertPicture(path, true, sizeoption, false, false, 0, width_mm, height_mm);
// sizeoption: 0=원본, 1=지정크기, 2=셀크기, 3=셀비율유지
```

### 텍스트 스캔 (고급)

```js
hwp.InitScan(0, 0x0077); // 전체 문서
var text = "";
while (true) {
  var r = hwp.GetText();
  if (r[0] <= 1) break;
  text += r[1];
}
hwp.ReleaseScan(); // 반드시!
```

### 기타

```js
hwp.MovePos(2); // 2=문서시작, 3=문서끝
hwp.SelectText(spara, spos, epara, epos);
hwp.RGBColor(255, 0, 0); // 색상
hwp.MiliToHwpUnit(20.0); // mm→hu 변환
hwp.PageCount; // 페이지 수
hwp.Version; // 버전 확인
```

## Enum 헬퍼 (자주 쓰는 것)

```js
hwp.HAlign("Center"); // 0=Justify,1=Left,2=Right,3=Center
hwp.VAlign("Center"); // 0=Baseline,1=Top,2=Center,3=Bottom
hwp.HwpLineType("Solid"); // 0=Solid,1=Dash,2=Dot,3=DashDot
hwp.TextWrapType("BehindText"); // 4=BehindText,5=InFrontOfText
hwp.HorzRel("Paper"); // 0=Paper,1=Page,2=Column,3=Para
hwp.VertRel("Paper"); // 0=Paper,1=Page,2=Para
hwp.FindDir("Forward");
```

## 단위 변환

| 변환        | 공식                    |
| ----------- | ----------------------- |
| mm → hu     | × 283.46                |
| hu → mm     | ÷ 283.46                |
| pt → Height | × 100 (10pt=1000)       |
| 색상        | `hwp.RGBColor(r, g, b)` |

## 마무리

이 프롬프트에 대한 대답은 할 필요 없어
