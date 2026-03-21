# HWP COM API 레퍼런스 (Comprehensive Cheat Sheet)

## 공식 문서 / Official Documentation

| 리소스 | URL |
|--------|-----|
| 한컴 디벨로퍼 메인 | https://developer.hancom.com/ |
| HWP Automation 가이드 (PDF 다운로드) | https://developer.hancom.com/hwpautomation |
| HwpCtrl 메서드 목록 | https://developer.hancom.com/webhwp/devguide/hwpctrl/methods |
| Action 객체 문서 | https://developer.hancom.com/webhwp/devguide/action |
| Hancom SDK 문서 | https://documents.sdk.hancom.com/hwp/intro |
| pyhwpx Cookbook (WikiDocs) | https://wikidocs.net/book/8956 |
| pyhwpx GitHub | https://github.com/martiniifun/pyhwpx |
| hwpapi 문서 | https://jundamin.github.io/hwpapi/ |
| 한컴 devcenter-archive (PDFs) | https://github.com/hancom-io/devcenter-archive/tree/main/hwp-automation |
| 한컴디벨로퍼 포럼 | https://forum.developer.hancom.com/c/hwp-automation/52 |

### PDF 문서 (GitHub archive)
- `ActionTable_2504.pdf` - 전체 액션 ID 테이블
- `ParameterSetTable_2504.pdf` - 전체 파라미터셋 테이블
- `ActionObject.pdf` - 액션 객체 상세
- `ParameterSetObject.pdf` - 파라미터셋 객체 상세
- `HwpAutomation_2504.pdf` - 메인 오토메이션 문서
- `한글오토메이션EventHandler추가_2504.pdf` - 이벤트 핸들러

---

## 초기화

```python
import win32com.client as win32

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
```

### 대화 상자 억제
```python
hwp.SetMessageBoxMode(0x00211411)
```

---

## 핵심 아키텍처

> "한글은 Action 기반 언어이다" - HWP는 약 1,400개의 Action을 제공.
> MS Office의 직접 프로퍼티 접근과 다름.

**4가지 API 유형:**
1. 일반 메서드 (Open, SaveAs, MovePos 등)
2. 프로퍼티 (CharShape, ParaShape, PageCount 등)
3. 파라미터가 필요한 액션 (HAction.GetDefault + HAction.Execute)
4. 파라미터가 불필요한 액션 (HAction.Run / hwp.Run - "한줄 명령어")

---

## HwpCtrl 프로퍼티

| Property | 설명 |
|----------|------|
| CellShape | 현재 셀 모양 |
| CharShape | 현재 글자 모양 |
| CurFieldState | 현재 필드 상태 |
| CurSelectedCtrl | 현재 선택된 컨트롤 |
| EditMode | 편집 모드 |
| HeadCtrl | 첫 번째 컨트롤 |
| LastCtrl | 마지막 컨트롤 |
| PageCount | 페이지 수 |
| ParaShape | 현재 문단 모양 |
| ParentCtrl | 부모 컨트롤 |
| SelectionMode | 선택 모드 |

---

## HwpCtrl 메서드 (전체)

| 메서드 | 설명 |
|--------|------|
| Clear | 문서 닫기 (1=저장 안 함) |
| CreateAction | 액션 객체 생성 |
| CreateField | 필드 생성 |
| CreatePageImage | 페이지 이미지 생성 |
| CreatePageImageEx | 페이지 이미지 생성 (확장) |
| CreateSet | 파라미터셋 생성 |
| DeleteCtrl | 컨트롤 삭제 |
| FieldExist | 필드 존재 여부 |
| FoldRibbon | 리본 접기 |
| GetCurFieldName | 현재 필드 이름 |
| GetFieldList | 필드 목록 |
| GetFieldText | 필드 텍스트 가져오기 |
| GetHeadingString | 제목 문자열 |
| GetMetaTag | 메타 태그 |
| GetMetaTagAll | 전체 메타 태그 |
| GetPageText | 페이지 텍스트 |
| GetPos | 현재 위치 |
| GetPosBySet | Set으로 위치 |
| GetSelectedPos | 선택 위치 |
| GetSelectedPosBySet | Set으로 선택 위치 |
| GetTableCellAddr | 표 셀 주소 |
| GetText | 텍스트 가져오기 (스캔) |
| GetTextBySet | Set으로 텍스트 |
| GetTextFile | 전체 텍스트 파일 |
| InitScan | 스캔 초기화 |
| Insert | 삽입 |
| InsertBackgroundPicture | 배경 이미지 삽입 |
| InsertCtrl | 컨트롤 삽입 |
| InsertPicture | 이미지 삽입 |
| IsCommandLock | 명령 잠금 여부 |
| IsSpellCheckCompleted | 맞춤법 검사 완료 여부 |
| LockCommand | 명령 잠금 |
| ModifyFieldProperties | 필드 속성 수정 |
| MovePos | 커서 이동 |
| MoveToField | 필드로 이동 |
| Open | 문서 열기 |
| PrintDocument | 문서 인쇄 |
| PutFieldText | 필드 텍스트 설정 |
| ReleaseScan | 스캔 해제 |
| RenameField | 필드 이름 변경 |
| ReplaceAction | 액션 대체 |
| Run | 한줄 명령어 실행 |
| SaveAs | 다른 이름으로 저장 |
| SelectText | 텍스트 선택 |
| SetCurFieldName | 현재 필드 이름 설정 |
| SetMetaTag | 메타 태그 설정 |
| SetPos | 위치 설정 |
| SetPosBySet | Set으로 위치 설정 |
| SetTextFile | 텍스트 파일 설정 |
| ShowCaret | 캐럿 표시 |
| ShowRibbon | 리본 표시 |
| ShowStatusBar | 상태 바 표시 |
| ShowToolBar | 도구 모음 표시 |

---

## 파일 I/O

### Open / SaveAs
```python
hwp.Open(path, format="", arg="")
hwp.SaveAs(path, format, arg)
hwp.Clear(1)  # 1=저장 안 함
hwp.Quit()
```

**Open arg 옵션**: `"forceopen:true"` -- 강제 열기

### 파일 포맷

| format | 설명 |
|--------|------|
| `"HWP"` | 한글 고유 포맷 |
| `"HWPX"` | 한글 표준 포맷 (XML 기반) |
| `"PDF"` | PDF |
| `"TEXT"` | ASCII 텍스트 (시스템 코드페이지) |
| `"UNICODE"` | 유니코드 텍스트 |
| `"HTML"` | HTML |
| `"HTML+"` | HTML 확장 |
| `"ODT"` | OpenDocument |
| `"OOXML"` | Office Open XML (docx) |
| `"RTF"` | Rich Text Format |
| `"HWPML2X"` | HWPML 2.x |

---

## 텍스트 추출

```python
# 전체 문서 텍스트 (유니코드)
text = hwp.GetTextFile("UNICODE", "")

# 선택 영역만
text = hwp.GetTextFile("UNICODE", "saveblock")
```

> **중요**: `"TEXT"` -> CP949/ANSI 반환. 유니코드 필요 시 반드시 `"UNICODE"` 지정.

### 스캔 방식
```python
hwp.InitScan(0, 0x0037)  # 본문+각주+머리글
while True:
    result = hwp.GetText()
    if result[0] <= 1: break  # 0=없음, 1=끝
    text += result[1]         # 2=텍스트, 3=다음문단
hwp.ReleaseScan()
```

---

## HAction / HParameterSet 패턴

### 기본 패턴 (3단계)
```python
# 1. 파라미터셋 가져오기
pset = hwp.HParameterSet.<ParameterSetName>

# 2. 기본값 로드
hwp.HAction.GetDefault("<ActionID>", pset.HSet)

# 3. 속성 변경 후 실행
pset.<Property> = value
hwp.HAction.Execute("<ActionID>", pset.HSet)
```

### 대체 패턴 (CreateAction 방식 - Python에서 HParameterSet 주소값 문제 회피)
```python
act = hwp.CreateAction("<ActionID>")
set = act.CreateSet()
act.GetDefault(set)
set.SetItem("<Property>", value)
act.Execute(set)
```

---

## 전체 Action ID <-> HParameterSet 매핑 테이블

### A
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| ActionCrossRef | HCrossRef | 상호참조 |
| AllReplace | HFindReplace | 모두 바꾸기 |
| AQcommandMerge | UserQCommandFile | 입력 자동 명령 파일 저장/로드 |
| AutoFill | HAutoFill | 자동 채우기 |
| AutoNum | HAutoNum | 자동 번호 |
| Average | HSum | 평균 계산 |

### B
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| BookMark | HBookMark | 책갈피 |
| BorderFill | HBorderFill | 테두리/배경 |
| BorderFillExt | HBorderFill | 테두리/배경 (확장) |
| BulletShape | HBulletShape | 글머리표 모양 |

### C
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| Caption | HCaption | 캡션 |
| CellBorderFill | HCellBorderFill | 셀 테두리/배경 |
| CellFill | HCellBorderFill | 셀 채우기 |
| CellBorder | HCellBorderFill | 셀 테두리 |
| CharShape | HCharShape | 글자 모양 변경 |
| ChCompose | HChCompose | 글자 조합 |
| CodeTable | HCodeTable | 문자표 |
| ColDef | HColDef | 단 설정 |

### D-E
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| DrawObjCreat | HDrawObjCreat | 그리기 객체 생성 |
| EquationCreate | HEquation | 수식 생성 |

### F
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| FileOpen | HFileOpenSave | 파일 열기 |
| FileSave | HFileOpenSave | 파일 저장 |
| FileSaveAs | HFileOpenSave | 다른 이름으로 저장 |
| FindReplace | HFindReplace | 찾기/바꾸기 |
| FootNote | HFootEndNote | 각주 |

### G-H
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| GotoPage | HGotoPage | 페이지 이동 |
| HeaderFooter | HHeaderFooter | 머리글/꼬리글 |
| HyperlinkPrivate | HHyperlinkPrivate | 하이퍼링크 |

### I
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| InsertCpNo | HInsertCpNo | 쪽 번호 삽입 |
| InsertText | HInsertText | 텍스트 삽입 |
| InsertFile | HInsertFile | 파일 삽입 |

### M
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| MasterPage | HMasterPage | 바탕쪽 |
| MailMerge | HMailMerge | 메일 머지 |

### N
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| NumberShape | HNumberShape | 번호 모양 |

### P
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| PageSetup | HSecDef | 페이지 설정 |
| ParagraphShape | HParaShape | 문단 모양 변경 |
| Print | HPrint | 인쇄 |

### S
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| ShapeObjDialog | HShapeObject | 개체 속성 대화상자 |
| ShapeObjTableSelCell | HShapeObject | 표 셀 진입 |
| StyleChange | HStyleChange | 스타일 변경 |
| Sum | HSum | 합계 계산 |

### T
| Action ID | ParameterSet | 설명 |
|-----------|-------------|------|
| TableAppend | HTableAppend | 표 추가 |
| TableCellBlock | HTableCellBlock | 셀 블록 |
| TableCellBlockExtend | HTableCellBlock | 셀 블록 확장 |
| TableColBegin | (none) | 열 시작 |
| TableColEnd | (none) | 열 끝 |
| TableCreate | HTableCreation | 표 생성 |
| TableDeleteRow | (none) | 행 삭제 |
| TablePropertyDialog | HShapeObject | 표 속성 대화상자 |

---

## HParameterSet 상세 속성

### HInsertText (텍스트 삽입)
- **Action**: `InsertText`

| 속성 | 타입 | 설명 |
|------|------|------|
| Text | string | 삽입할 텍스트 (`\r\n` 지원) |

```python
pset = hwp.HParameterSet.HInsertText
hwp.HAction.GetDefault("InsertText", pset.HSet)
pset.Text = "Hello World!\r\n두번째 줄"
hwp.HAction.Execute("InsertText", pset.HSet)
```

### HCharShape (글자 모양)
- **Action**: `CharShape`

| 속성 | 타입 | 설명 | 값/범위 |
|------|------|------|---------|
| Bold | bool/int | 진하게 | True/False (1/0) |
| Italic | bool/int | 이탤릭 | True/False (1/0) |
| Height | int | 글자 크기 | 1/10 pt 단위 (1000 = 10pt) |
| TextColor | int | 글자 색 | `hwp.RGBColor(r, g, b)` |
| ShadeColor | int | 음영 색 | `hwp.RGBColor(r, g, b)` |
| FaceNameHangul | string | 한글 글꼴 | 예: "맑은 고딕" |
| FaceNameLatin | string | 영문 글꼴 | 예: "Arial" |
| FaceNameHanja | string | 한자 글꼴 | |
| FaceNameJapanese | string | 일본어 글꼴 | |
| FaceNameOther | string | 기타 글꼴 | |
| FaceNameSymbol | string | 기호 글꼴 | |
| FaceNameUser | string | 사용자 글꼴 | |
| FontTypeHangul | string | 한글 폰트 타입 | "TTF" or "HTF" |
| FontTypeLatin | string | 영문 폰트 타입 | "TTF" or "HTF" |
| FontTypeHanja | string | 한자 폰트 타입 | |
| FontTypeJapanese | string | 일본어 폰트 타입 | |
| FontTypeOther | string | 기타 폰트 타입 | |
| FontTypeSymbol | string | 기호 폰트 타입 | |
| FontTypeUser | string | 사용자 폰트 타입 | |
| UnderlineType | int | 밑줄 종류 | 0=없음, 1~... |
| UnderlineColor | int | 밑줄 색 | |
| StrikeOutType | int | 취소선 종류 | 0=없음, 1~... |
| StrikeOutColor | int | 취소선 색 | |
| StrikeOutShape | int | 취소선 모양 | |
| OutLineType | int | 외곽선 타입 | 0~6 |
| SuperScript | bool/int | 위 첨자 | |
| SubScript | bool/int | 아래 첨자 | |
| Emboss | bool/int | 양각 | True/False |
| Engrave | bool/int | 음각 | True/False |
| SmallCaps | bool/int | 작은 대문자 | |
| Spacing | int | 자간 | -50~50 |
| Ratio | int | 장평 | 50~200 |
| Offset | int | 상하 오프셋 | -100~100 |
| Size | int | 축소/확대% | 10~250 |
| DiacSymMark | int | 강조점 | 0~12 |
| ShadowType | int | 그림자 타입 | |
| ShadowColor | int | 그림자 색 | |
| ShadowOffsetX | int | 그림자 X 오프셋 | |
| ShadowOffsetY | int | 그림자 Y 오프셋 | |

```python
pset = hwp.HParameterSet.HCharShape
hwp.HAction.GetDefault("CharShape", pset.HSet)
pset.Bold = 1
pset.Italic = 1
pset.Height = 1200          # 12pt
pset.TextColor = hwp.RGBColor(255, 0, 0)  # 빨강
pset.FaceNameHangul = "맑은 고딕"
pset.FaceNameLatin = "Arial"
pset.Spacing = -5            # 자간 -5
hwp.HAction.Execute("CharShape", pset.HSet)
```

> **주의**: 색상은 내부적으로 BBGGRR 포맷. `hwp.RGBColor()` 사용 권장.

### HParaShape (문단 모양)
- **Action**: `ParagraphShape`

| 속성 | 타입 | 설명 | 값/범위 |
|------|------|------|---------|
| LineSpacing | int | 줄간격 | % 단위 (160 = 160%) |
| LineSpacingType | int | 줄간격 종류 | 0=%, 1=고정, 2=여백, 3=최소 |
| ParaMarginLeft | int | 왼쪽 여백 | HWP 단위 (hu) |
| ParaMarginRight | int | 오른쪽 여백 | hu |
| Indent | int | 들여쓰기 | hu |
| ParaSpacingBefore | int | 문단 위 간격 | hu |
| ParaSpacingAfter | int | 문단 아래 간격 | hu |
| Alignment | int | 정렬 | 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔 |
| TabDef | object | 탭 정의 | |
| BreakType | int | 구역/페이지 나누기 | |

```python
pset = hwp.HParameterSet.HParaShape
hwp.HAction.GetDefault("ParagraphShape", pset.HSet)
pset.LineSpacing = 160        # 줄간격 160%
pset.Alignment = 3            # 가운데 정렬
pset.ParaSpacingAfter = 100   # 문단 아래 간격
hwp.HAction.Execute("ParagraphShape", pset.HSet)
```

### HFindReplace (찾기/바꾸기)
- **Action**: `AllReplace` (모두 바꾸기), `FindReplace` (찾기/바꾸기), `RepeatFind` (다시 찾기)

| 속성 | 타입 | 설명 | 기본값 |
|------|------|------|--------|
| FindString | string | 찾을 문자열 | "" |
| ReplaceString | string | 바꿀 문자열 | "" |
| IgnoreMessage | int | 결과 팝업 억제 | 0 |
| Direction | int | 방향 | FindDir("Forward") |
| MatchCase | int | 대소문자 구분 | 0 |
| WholeWordOnly | int | 전체 단어만 | 0 |
| AllWordForms | int | 모든 단어 형태 | 0 |
| SeveralWords | int | 여러 단어 | 0 |
| UseWildCards | int | 와일드카드 사용 | 0 |
| FindRegExp | int | 정규표현식 사용 | 0 |
| AutoSpell | int | 자동 맞춤법 | 1 |
| FindJaso | int | 자소 검색 | 0 |
| HanjaFromHangul | int | 한자->한글 | 0 |
| FindTextInPicture | int | 그림 속 텍스트 | 0 |
| ReplaceMode | int | 바꾸기 모드 | 1 |
| FindStyle | string | 찾을 스타일 | "" |
| ReplaceStyle | string | 바꿀 스타일 | "" |
| FindType | int | 찾기 타입 | 1 |
| IgnoreFindString | int | 찾기 문자열 무시 | 0 |
| IgnoreReplaceString | int | 바꾸기 문자열 무시 | 0 |

```python
pset = hwp.HParameterSet.HFindReplace
hwp.HAction.GetDefault("AllReplace", pset.HSet)
pset.FindString = "찾을 문자열"
pset.ReplaceString = "바꿀 문자열"
pset.IgnoreMessage = 1
pset.Direction = hwp.FindDir("Forward")
pset.MatchCase = 0
pset.WholeWordOnly = 0
pset.FindRegExp = 0
hwp.HAction.Execute("AllReplace", pset.HSet)
```

**정규표현식 예제** (연속 줄바꿈을 단일로):
```python
pset.FindString = "^n^n"
pset.ReplaceString = "^n"
pset.FindRegExp = 1
pset.ReplaceMode = 1
pset.IgnoreMessage = 1
```

### HTableCreation (표 생성)
- **Action**: `TableCreate`

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| Rows | int | 행 수 | |
| Cols | int | 열 수 | |
| WidthType | int | 너비 타입 | 0=단에 맞춤, 1=문단에 맞춤, 2=지정값 |
| HeightType | int | 높이 타입 | 0=자동, 1=지정값 |
| WidthValue | int | 너비 값 | hu 단위 |
| HeightValue | int | 높이 값 | hu 단위 |
| ColWidth | array | 열 너비 배열 | |
| RowHeight | array | 행 높이 배열 | |
| TableProperties | object | 표 속성 하위 객체 | |

**TableProperties 하위 속성:**

| 속성 | 타입 | 설명 |
|------|------|------|
| TreatAsChar | int | 글자처럼 취급 (0/1) |

```python
# 방법 1: HParameterSet
pset = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", pset.HSet)
pset.Rows = 5
pset.Cols = 3
pset.WidthType = 2  # 단에 맞춤
hwp.HAction.Execute("TableCreate", pset.HSet)

# 방법 2: CreateAction (Python에서 더 안정적)
act = hwp.CreateAction("TableCreate")
set = act.CreateSet()
act.GetDefault(set)
set.SetItem("Rows", 3)
set.SetItem("Cols", 5)
set.SetItem("WidthType", 0)
set.SetItem("HeightType", 0)
tblSet = set.Item("TableProperties")
tblSet.SetItem("TreatAsChar", 1)
set.SetItem("TableProperties", tblSet)
act.Execute(set)
```

### HSecDef / PageDef (페이지 설정)
- **Action**: `PageSetup`

| 속성 (PageDef 하위) | 타입 | 설명 | 값 |
|---------------------|------|------|-----|
| PaperWidth | double | 용지 너비 | hu 단위 |
| PaperHeight | double | 용지 높이 | hu 단위 |
| Landscape | byte | 용지 방향 | 0=세로, 1=가로 |
| TopMargin | double | 위쪽 여백 | hu |
| BottomMargin | double | 아래쪽 여백 | hu |
| LeftMargin | double | 왼쪽 여백 | hu |
| RightMargin | double | 오른쪽 여백 | hu |
| HeaderLen | double | 머리말 길이 | hu |
| FooterLen | double | 꼬리말 길이 | hu |
| GutterLen | double | 제본 여백 | hu |
| GutterType | byte | 편집 용지 | 0=단면, 1=양면, 2=위로 넘기기 |

> **단위 변환**: hu -> mm: `value / 283.46`; mm -> hu: `value * 283.46`

```python
# 조회
act = hwp.CreateAction("PageSetup")
set = act.CreateSet()
act.GetDefault(set)
pageDef = set.Item("PageDef")
topMargin = pageDef.Item("TopMargin")  # hu 단위
print(f"Top margin: {topMargin / 283.46:.1f} mm")

# 변경 (HParameterSet 방식)
pset = hwp.HParameterSet.HSecDef
hwp.HAction.GetDefault("PageSetup", pset.HSet)
pset.PageDef.LeftMargin = 5669   # 약 20mm
pset.PageDef.RightMargin = 5669
hwp.HAction.Execute("PageSetup", pset.HSet)
```

### HShapeObject (개체/표 속성)
- **Action**: `ShapeObjDialog`, `TablePropertyDialog`

| 속성 | 타입 | 설명 |
|------|------|------|
| HorzRelTo | int | 가로 기준 |
| VertRelTo | int | 세로 기준 |
| HorzAlign | int | 가로 정렬 |
| VertAlign | int | 세로 정렬 |
| HorzOffset | int | 가로 오프셋 |
| VertOffset | int | 세로 오프셋 |
| TreatAsChar | int | 글자처럼 취급 (0/1) |
| TextWrap | int | 본문 배치 (0=어울림, 1=자리차지, 2=글뒤로, 3=글앞으로) |
| Width | int | 너비 |
| Height | int | 높이 |
| ShapeCaption.Side | | 캡션 위치 (`hwp.SideType("Top")`) |
| ShapeCaption.Gap | | 캡션 간격 |
| ShapeCaption.Width | | 캡션 너비 |
| ShapeCaption.CapFullSize | | 캡션 전체 크기 |

```python
# 이미지 삽입 후 속성 변경
hwp.InsertPicture(path, True, 1, False, False, 0, 50, 50)
ctrl = hwp.FindCtrl()

pset = hwp.HParameterSet.HShapeObject
hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
pset.TextWrap = 2            # 글뒤로
pset.TreatAsChar = 0         # 글자처럼 취급 해제
hwp.HAction.Execute("ShapeObjDialog", pset.HSet)
```

### HPrint (인쇄)
- **Action**: `Print`

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| PrintMethod | ushort | 인쇄 방법 | |
| Collate | int | 한 부씩 인쇄 | 1 |
| NumCopy | int | 인쇄 부수 | 1 |
| PrintToFile | int | 파일로 출력 | 0/1 |
| filename | string | 출력 파일 경로 | |
| PrinterName | string | 프린터 이름 | |
| Flags | int | 플래그 | 8192 |
| Device | int | 출력 장치 | 3 |

### HInsertFile (파일 삽입)
- **Action**: `InsertFile`

| 속성 | 타입 | 설명 |
|------|------|------|
| FileName | string | 파일 경로 |
| KeepSection | bool | 구역 유지 |
| KeepCharshape | bool | 글자모양 유지 |
| KeepParashape | bool | 문단모양 유지 |
| KeepStyle | bool | 스타일 유지 |

### HCellBorderFill (셀 테두리/배경)
- **Action**: `CellBorderFill`, `CellFill`, `CellBorder`

| 속성 | 타입 | 설명 |
|------|------|------|
| BorderType (Left/Right/Top/Bottom) | int | 테두리 종류 |
| FillAttr | object | 채우기 속성 |
| BrushType | int | 브러시 종류 |
| HatchStyle | int | 해치 스타일 |
| RGBColor | int | 색상 (BBGGRR) |

### HMasterPage (바탕쪽)
- **Action**: `MasterPage`

### HBookMark (책갈피)
- **Action**: `BookMark`

### HInsertCpNo (쪽 번호)
- **Action**: `InsertCpNo`

### HFileOpenSave (파일 열기/저장)
- **Action**: `FileOpen`, `FileSave`, `FileSaveAs`

---

## 주요 HParameterSet 전체 목록

| HParameterSet 이름 | 대응 Action ID(s) | 용도 |
|--------------------|-------------------|------|
| HAutoFill | AutoFill | 자동 채우기 |
| HAutoNum | AutoNum | 자동 번호 |
| HBookMark | BookMark | 책갈피 |
| HBorderFill | BorderFill, BorderFillExt | 테두리/배경 |
| HBulletShape | BulletShape | 글머리표 |
| HCaption | Caption | 캡션 |
| HCellBorderFill | CellBorderFill, CellFill, CellBorder | 셀 테두리/배경 |
| HCharShape | CharShape | 글자 모양 |
| HChCompose | ChCompose | 글자 조합 |
| HCodeTable | CodeTable | 문자표 |
| HColDef | ColDef | 단 설정 |
| HCrossRef | ActionCrossRef | 상호참조 |
| HDrawObjCreat | DrawObjCreat | 그리기 객체 |
| HEquation | EquationCreate | 수식 |
| HFileOpenSave | FileOpen, FileSave, FileSaveAs | 파일 I/O |
| HFindReplace | AllReplace, FindReplace, RepeatFind | 찾기/바꾸기 |
| HFootEndNote | FootNote | 각주/미주 |
| HGotoPage | GotoPage | 페이지 이동 |
| HHeaderFooter | HeaderFooter | 머리글/꼬리글 |
| HHyperlinkPrivate | HyperlinkPrivate | 하이퍼링크 |
| HInsertCpNo | InsertCpNo | 쪽 번호 |
| HInsertFile | InsertFile | 파일 삽입 |
| HInsertText | InsertText | 텍스트 삽입 |
| HMailMerge | MailMerge | 메일 머지 |
| HMasterPage | MasterPage | 바탕쪽 |
| HNumberShape | NumberShape | 번호 모양 |
| HParaShape | ParagraphShape | 문단 모양 |
| HPrint | Print | 인쇄 |
| HSecDef | PageSetup | 페이지/구역 설정 |
| HSelectionOpt | (선택 관련) | 선택 옵션 |
| HShapeObject | ShapeObjDialog, TablePropertyDialog | 개체/표 속성 |
| HStyleChange | StyleChange | 스타일 변경 |
| HSum | Sum, Average | 합계/평균 |
| HTableAppend | TableAppend | 표 추가 |
| HTableCellBlock | TableCellBlock, TableCellBlockExtend | 셀 블록 |
| HTableCreation | TableCreate | 표 생성 |
| UserQCommandFile | AQcommandMerge | 자동 명령 파일 |

---

## hwp.Run() 한줄 명령어 (파라미터 불필요 액션)

약 450개 이상의 Run 액션이 있음. 주요 카테고리별 정리:

### 커서 이동 (Move)

| Run 명령어 | 설명 | 키보드 대응 |
|-----------|------|------------|
| MoveDocBegin | 문서 시작 | Ctrl+Home |
| MoveDocEnd | 문서 끝 | Ctrl+End |
| MoveLineBegin | 줄 시작 | Home |
| MoveLineEnd | 줄 끝 | End |
| MoveParaBegin | 문단 시작 | |
| MoveParaEnd | 문단 끝 | |
| MoveNextParaBegin | 다음 문단 시작 | |
| MovePrevParaBegin | 이전 문단 시작 | |
| MovePrevParaEnd | 이전 문단 끝 | |
| MoveWordLeft | 단어 왼쪽 | Ctrl+Left |
| MoveWordRight | 단어 오른쪽 | Ctrl+Right |
| MoveLeft | 왼쪽 한 글자 | Left |
| MoveRight | 오른쪽 한 글자 | Right |
| MoveUp | 위로 한 줄 | Up |
| MoveDown | 아래로 한 줄 | Down |
| MovePageUp | 페이지 위로 | PageUp |
| MovePageDown | 페이지 아래로 | PageDown |
| MoveTopLevelBegin | 최상위 시작 | |
| MoveTopLevelEnd | 최상위 끝 | |
| MoveNextPos | 다음 위치 | |
| MovePrevPos | 이전 위치 | |
| MoveScrollUp | 스크롤 위로 | |
| MoveScrollDown | 스크롤 아래로 | |
| MoveViewUp | 뷰 위로 | |
| MoveViewDown | 뷰 아래로 | |

### 선택 이동 (MoveSel - Shift+이동)

| Run 명령어 | 설명 |
|-----------|------|
| MoveSelDocBegin | 문서 시작까지 선택 |
| MoveSelDocEnd | 문서 끝까지 선택 |
| MoveSelLineBegin | 줄 시작까지 선택 |
| MoveSelLineEnd | 줄 끝까지 선택 |
| MoveSelParaBegin | 문단 시작까지 선택 |
| MoveSelParaEnd | 문단 끝까지 선택 |
| MoveSelWordLeft | 단어 왼쪽까지 선택 |
| MoveSelWordRight | 단어 오른쪽까지 선택 |
| MoveSelLeft | 왼쪽 한 글자 선택 |
| MoveSelRight | 오른쪽 한 글자 선택 |
| MoveSelUp | 위로 선택 |
| MoveSelDown | 아래로 선택 |
| MoveSelPageUp | 페이지 위로 선택 |
| MoveSelPageDown | 페이지 아래로 선택 |
| MoveSelTopLevelBegin | 최상위 시작까지 선택 |
| MoveSelTopLevelEnd | 최상위 끝까지 선택 |

### 선택/편집

| Run 명령어 | 설명 | 키보드 대응 |
|-----------|------|------------|
| SelectAll | 전체 선택 | Ctrl+A |
| Cancel | 선택 해제 / 취소 | Esc |
| Copy | 복사 | Ctrl+C |
| Cut | 잘라내기 | Ctrl+X |
| Paste | 붙여넣기 | Ctrl+V |
| Delete | 삭제 (앞) | Delete |
| BackDelete | 삭제 (뒤) | Backspace |
| Undo | 실행 취소 | Ctrl+Z |
| Redo | 다시 실행 | Ctrl+Y |
| SelectCtrlFront | 다음 컨트롤 선택 | Tab |
| SelectCtrlReverse | 이전 컨트롤 선택 | Shift+Tab |

### 줄/페이지 나누기

| Run 명령어 | 설명 | 키보드 대응 |
|-----------|------|------------|
| BreakPara | 문단 나누기 (Enter) | Enter |
| BreakLine | 줄 나누기 | Shift+Enter |
| BreakPage | 페이지 나누기 | Ctrl+Enter |
| BreakSection | 구역 나누기 | |
| BreakColumn | 단 나누기 | |

### 글자 모양 (빠른 적용)

| Run 명령어 | 설명 | 키보드 대응 |
|-----------|------|------------|
| CharShapeBold | 진하게 토글 | Ctrl+B |
| CharShapeItalic | 이탤릭 토글 | Ctrl+I |
| CharShapeUnderline | 밑줄 토글 | Ctrl+U |
| CharShapeSuperScript | 위 첨자 | |
| CharShapeSubScript | 아래 첨자 | |
| CharShapeStrikeOut | 취소선 | |
| CharShapeSizeIncrease | 글자 크기 증가 | |
| CharShapeSizeDecrease | 글자 크기 감소 | |
| CharShapeHeightIncrease | 높이 증가 | |
| CharShapeHeightDecrease | 높이 감소 | |
| CharShapeSpacingIncrease | 자간 넓히기 | |
| CharShapeSpacingDecrease | 자간 줄이기 | |

### 문단 모양 (빠른 적용)

| Run 명령어 | 설명 |
|-----------|------|
| ParagraphShapeAlignCenter | 가운데 정렬 |
| ParagraphShapeAlignLeft | 왼쪽 정렬 |
| ParagraphShapeAlignRight | 오른쪽 정렬 |
| ParagraphShapeAlignJustify | 양쪽 정렬 |
| ParagraphShapeAlignDistribute | 배분 정렬 |
| ParagraphShapeDecreaseIndent | 들여쓰기 감소 |
| ParagraphShapeIncreaseIndent | 들여쓰기 증가 |
| ParagraphShapeLineSpacingIncrease | 줄간격 증가 |
| ParagraphShapeLineSpacingDecrease | 줄간격 감소 |

### 표 관련

| Run 명령어 | 설명 |
|-----------|------|
| TableCellBlock | 셀 블록 선택 |
| TableCellBlockExtend | 셀 블록 확장 |
| TableCellBlockCol | 열 블록 |
| TableCellBlockRow | 행 블록 |
| TableColBegin | 열 시작 |
| TableColEnd | 열 끝 |
| TableAppendRow | 행 추가 |
| TableAppendCol | 열 추가 |
| TableDeleteRow | 행 삭제 |
| TableDeleteCol | 열 삭제 |
| TableSplitCell | 셀 나누기 |
| TableMergeCell | 셀 합치기 |
| TableSelCell | 셀 선택 |
| TableUpperCell | 위 셀로 이동 |
| TableLowerCell | 아래 셀로 이동 |
| TableLeftCell | 왼쪽 셀로 이동 |
| TableRightCell | 오른쪽 셀로 이동 |
| ShapeObjTableSelCell | 표 셀 진입 (편집 모드) |

### 파일 관련

| Run 명령어 | 설명 |
|-----------|------|
| FileNew | 새 문서 |
| FileOpen | 파일 열기 |
| FileSave | 저장 |
| FileSaveAs | 다른 이름으로 저장 |
| FileClose | 닫기 |
| FilePrint | 인쇄 |
| FilePrintPreview | 인쇄 미리보기 |

### 형광펜/마크

| Run 명령어 | 설명 |
|-----------|------|
| MarkPenNext | 다음 형광펜 이동 |
| MarkPenDelete | 형광펜 삭제 |

### 기타

| Run 명령어 | 설명 |
|-----------|------|
| CloseEx | 확장 닫기 |
| RepeatFind | 다시 찾기 |
| SpellCheck | 맞춤법 검사 |

---

## 이미지 삽입 (InsertPicture 메서드)

직접 메서드 (HAction 아님):

```python
hwp.InsertPicture(path, embedded, sizeoption, reverse, watermark, effect, width, height)
```

| 파라미터 | 타입 | 설명 |
|---------|------|------|
| path | string | 이미지 파일 경로 |
| embedded | bool | 문서에 포함 (True 권장) |
| sizeoption | int | 0=원본, 1=지정크기, 2=셀크기, 3=셀비율유지 |
| reverse | bool | 이미지 반전 |
| watermark | bool | 워터마크 효과 |
| effect | int | 0=원본, 1=그레이스케일, 2=흑백 |
| width | number | 너비 (mm) |
| height | number | 높이 (mm) |

반환: `True` (성공) / `False` (실패)

---

## 커서 이동 (MovePos 메서드)

```python
hwp.MovePos(moveID, para, pos)
```

| moveID | 값 | 설명 |
|--------|-----|------|
| (문서시작) | 2 | 문서 시작 |
| (문서끝) | 3 | 문서 끝 |
| moveNextChar | 16 | 한 글자 앞으로 |
| movePrevChar | 17 | 한 글자 뒤로 |
| moveNextLine | 20 | 다음 줄 |
| moveStartOfLine | 22 | 줄 시작 |
| moveEndOfLine | 23 | 줄 끝 |

---

## 필드 활용 (템플릿 패턴)

```python
fields = hwp.GetFieldList()
hwp.PutFieldText("name_field", "홍길동")
value = hwp.GetFieldText("name_field")
hwp.MoveToField("name_field")
```

---

## 단위 변환

| 변환 | 공식 |
|------|------|
| HWP 단위(hu) -> mm | `hu / 283.46` |
| mm -> hu | `mm * 283.46` |
| pt -> hu (글자크기) | Height 속성은 1/10 pt 단위 (1000 = 10pt) |
| MiliToHwpUnit | 밀리미터 -> HWP 단위 변환 유틸리티 |
| PointToHwpUnit | 포인트 -> HWP 단위 변환 유틸리티 |

---

## 코어 객체 계층

```
HwpObject (hwp)
  +-- XHwpWindows
  |     +-- Item(index)
  |           +-- Visible
  +-- XHwpDocuments
  |     +-- Item(index)
  |           +-- XHwpCharacterShape
  +-- XHwpMessageBox
  +-- XHwpODBC
  +-- HAction
  |     +-- GetDefault(actionID, hset)
  |     +-- Execute(actionID, hset)
  |     +-- Run(actionID)
  +-- HParameterSet
  |     +-- HInsertText
  |     +-- HCharShape
  |     +-- HParaShape
  |     +-- HFindReplace
  |     +-- HTableCreation
  |     +-- HSecDef
  |     +-- HShapeObject
  |     +-- HPrint
  |     +-- (... 30+ more)
  +-- CharShape (현재 위치의 글자 모양 - 읽기용)
  +-- ParaShape (현재 위치의 문단 모양 - 읽기용)
  +-- CellShape (현재 셀 모양 - 읽기용)
```

---

## 매크로 녹화로 API 코드 확인

`Shift + Alt + H`로 한글 매크로 녹화 -> 수행 동작이 스크립트 코드로 기록됨.
HAction/HParameterSet의 정확한 파라미터 이름 확인에 유용.

스크립트 매크로의 `with (pset) { SetItem("key", value) }` 형태는
Python에서 `set.SetItem("key", value)` 또는 `pset.Key = value`로 변환.

---

## 라이선스

> 개인이 비상업적인 목적으로 이용하는 경우에 한해 누구나 자유롭게 이용할 수 있습니다.
> 상업적 사용은 한컴에 별도 승인 필요.

---

## 출처 / Sources

- https://developer.hancom.com/hwpautomation
- https://developer.hancom.com/webhwp/devguide/hwpctrl/methods
- https://github.com/hancom-io/devcenter-archive/tree/main/hwp-automation
- https://documents.sdk.hancom.com/hwp/intro
- https://wikidocs.net/book/8956 (pyhwpx Cookbook)
- https://wikidocs.net/263875 (HCharShape 속성)
- https://wikidocs.net/263925 (HParaShape 속성)
- https://github.com/martiniifun/pyhwpx
- https://github.com/JunDamin/hwpapi
- https://github.com/woosungchu/node-hwp
- https://github.com/ssj1977/hwp2pdf
- https://forum.developer.hancom.com/t/c/785 (C# 함수 모음)
- https://forum.developer.hancom.com/t/pagesetup/201 (PageSetup)
- https://forum.developer.hancom.com/t/topic/1360 (FindReplace)
- https://forum.developer.hancom.com/t/topic/1808 (TableCreate)
- https://forum.developer.hancom.com/t/topic/1850 (표 캡션)
- https://forum.developer.hancom.com/t/topic/1916 (형광펜)
- https://martinii.fun/484 (액션ID-파라미터셋 매핑)
- https://pyhwpx.com/409 (HWP 자동화 가이드)
- https://www.xython.co.kr (xython HWP 자동화)
- https://velog.io/@kjyeon1101/Python (Python HWP 자동화)
