# HWP COM API 레퍼런스

## 공식 문서

| 리소스 | URL |
|--------|-----|
| 한컴 디벨로퍼 메인 | https://developer.hancom.com/ |
| HwpCtrl 메서드 목록 | https://developer.hancom.com/webhwp/devguide/hwpctrl/methods |
| Action 객체 문서 | https://developer.hancom.com/webhwp/devguide/action |
| pyhwpx Cookbook | https://wikidocs.net/book/8956 |
| hwpapi 문서 | https://jundamin.github.io/hwpapi/ |

## 초기화

```python
hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
```

## 주요 메서드

### 파일 I/O

| 메서드 | 시그니처 | 설명 |
|--------|---------|------|
| Open | `Open(path, format?, arg?)` | 문서 열기 |
| SaveAs | `SaveAs(path, format, arg)` | 다른 이름으로 저장 |
| Clear | `Clear(option)` | 문서 닫기 (1=저장 안 함) |

**Open arg 옵션**: `"forceopen:true"` — 강제 열기

### 파일 포맷 (Open/SaveAs 공통)

| format | 설명 |
|--------|------|
| `"HWP"` | 한글 고유 포맷 |
| `"HWPX"` | 한글 표준 포맷 (XML 기반) |
| `"PDF"` | PDF |
| `"TEXT"` | ASCII 텍스트 (시스템 코드페이지) |
| `"UNICODE"` | 유니코드 텍스트 |
| `"HTML"` | HTML |
| `"ODT"` | OpenDocument |

### 텍스트 추출

```python
# 전체 문서 텍스트 (유니코드)
text = hwp.GetTextFile("UNICODE", "")

# 선택 영역만
text = hwp.GetTextFile("UNICODE", "saveblock")
```

> **중요**: `"TEXT"` → CP949/ANSI 반환. 유니코드 필요 시 반드시 `"UNICODE"` 지정.

### 텍스트 추출 (스캔 방식)

```python
hwp.InitScan(0, 0x0037)  # 본문+각주+머리글
while True:
    result = hwp.GetText()
    if result[0] <= 1: break  # 0=없음, 1=끝
    text += result[1]         # 2=텍스트, 3=다음문단
hwp.ReleaseScan()
```

## HAction / HParameterSet 패턴

HWP의 핵심 프로그래밍 모델. MS Office의 직접 프로퍼티 접근과 다름.

### 텍스트 삽입

```python
pset = hwp.HParameterSet.HInsertText
hwp.HAction.GetDefault("InsertText", pset.HSet)
pset.Text = "Hello World!"
hwp.HAction.Execute("InsertText", pset.HSet)
```

### 찾기/바꾸기

```python
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
opt = hwp.HParameterSet.HFindReplace
opt.FindString = "찾을 문자열"
opt.ReplaceString = "바꿀 문자열"
opt.IgnoreMessage = 1    # 결과 팝업 억제
opt.Direction = 0        # 0=앞으로, 1=뒤로
opt.WholeWordOnly = 0
opt.UseRegExp = 0
hwp.HAction.Execute("AllReplace", opt.HSet)
```

### 글자 모양 변경

```python
pset = hwp.HParameterSet.HCharShape
hwp.HAction.GetDefault("CharShape", pset.HSet)
pset.TextColor = hwp.RGBColor(255, 0, 0)  # 빨간색
pset.Height = 1000    # 10pt (1/10 pt 단위)
pset.Bold = 1
hwp.HAction.Execute("CharShape", pset.HSet)
```

### 표 생성

```python
pset = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", pset.HSet)
pset.Rows = 5
pset.Cols = 3
pset.WidthType = 2  # 단에 맞춤
hwp.HAction.Execute("TableCreate", pset.HSet)
```

### 이미지 삽입

```python
hwp.InsertPicture(path, embedded, sizeoption, reverse, watermark, effect, width, height)
# sizeoption: 0=원본, 1=지정크기, 2=셀크기, 3=셀비율유지
# width/height: mm 단위
```

## 주요 HParameterSet 이름

| 이름 | 용도 |
|------|------|
| HInsertText | 텍스트 삽입 |
| HCharShape | 글자 모양 |
| HParaShape | 문단 모양 |
| HFindReplace | 찾기/바꾸기 |
| HFileOpenSave | 파일 열기/저장 |
| HSelectionOpt | 선택 옵션 |
| HTableCreation | 표 생성 |
| HPageDef | 페이지 설정 |

## 커서 이동

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

Run 명령으로도 가능:
```python
hwp.Run("MoveDocBegin")
hwp.Run("MoveDocEnd")
hwp.Run("MoveLineEnd")
```

## 필드 활용 (템플릿 패턴)

```python
fields = hwp.GetFieldList()
hwp.PutFieldText("name_field", "홍길동")
value = hwp.GetFieldText("name_field")
hwp.MoveToField("name_field")
```

## 매크로 녹화로 API 코드 확인

`Shift + Alt + H`로 한글 매크로 녹화 → 수행 동작이 스크립트 코드로 기록됨.
HAction/HParameterSet의 정확한 파라미터 이름 확인에 유용.

## 출처

- https://developer.hancom.com/webhwp/devguide/hwpctrl/methods
- https://wikidocs.net/book/8956
- https://github.com/martiniifun/pyhwpx
- https://jundamin.github.io/hwpapi/
