# HWP COM 고급 API 레퍼런스

## 필드 (누름틀) 자동화

### CreateField
```python
hwp.CreateField(direction, memo, name)
# direction: 안내문 (필드 비었을 때 표시)
# memo: 도움말 (데스크톱 전용)
# name: 필드 이름
```

### GetFieldList
```python
field_list = hwp.GetFieldList(number, option)
# number: 0=순차, 1=일련번호({{0}},{{1}}), 2=개수
# option: 0=전체, 1=셀, 2=누름틀, 4=선택영역 (OR 결합)
# 반환: 필드명들을 \x02로 구분한 문자열
fields = field_list.split("\x02")
```

### PutFieldText / GetFieldText
```python
hwp.PutFieldText("이름", "홍길동")           # "이름" 필드 전부 채움
hwp.PutFieldText("이름{{0}}", "홍길동")      # 첫 번째만
hwp.PutFieldText("a\x02b", "값1\x02값2")    # 여러 필드 한번에

text = hwp.GetFieldText("본문{{0}}")
```

> **주의**: `"fieldname"` vs `"fieldname{{0}}"` — 전자는 ALL 매칭, 후자는 특정 인스턴스만.

### MoveToField
```python
hwp.MoveToField(field, text=True, start=True, select=False)
# text: True=내부텍스트, False=필드코드
# start: True=시작, False=끝
# select: True=블록선택
```

---

## InsertPicture (이미지 삽입)

```python
ctrl = hwp.InsertPicture(path, embedded, sizeoption, reverse, watermark, effect, width, height)
```

| 파라미터 | 설명 |
|---------|------|
| sizeoption=0 | 원본 크기 |
| sizeoption=1 | 지정 크기 (width/height mm) |
| sizeoption=2 | 셀 크기 |
| sizeoption=3 | 셀 비율 유지 |
| effect=0 | 원본 |
| effect=1 | 그레이스케일 |
| effect=2 | 흑백 |

> `embedded`는 실질적으로 항상 True.

---

## InsertBackgroundPicture (배경 이미지)

```python
hwp.InsertBackgroundPicture(bordertype, path, embedded, filloption, watermark, effect, brightness, contrast)
# bordertype: "SelectedCell" (설정) / "SelectedCellDelete" (삭제)
# filloption: 0=바둑판, 5=크기맞춤, 6=가운데, 15=비율유지
```

---

## CreatePageImage (페이지 → 이미지)

```python
for i in range(hwp.PageCount):
    hwp.CreatePageImage(f"page_{i}.png", i, "png")
# pgno: 0-based
# format: "jpg", "jpeg", "png"
```

---

## InitScan / GetText / ReleaseScan (텍스트 스캔)

```python
hwp.InitScan(0, 0x0077)  # 전체 문서 (0x0070=문서시작 + 0x0007=문서끝)
full_text = ""
while True:
    state, text = hwp.GetText()
    if state in (0, 1):   # 0=없음, 1=끝
        break
    if state >= 101:       # 에러
        break
    full_text += text      # 2=텍스트, 3=다음문단
hwp.ReleaseScan()  # 반드시 호출!
```

> **GetText 반환 state**: 0=없음, 1=끝, 2=텍스트, 3=다음문단, 4=컨트롤진입, 5=컨트롤탈출, 101=미초기화, 102=변환실패

---

## SelectText / GetSelectedPos (선택 관리)

```python
hwp.SelectText(spara, spos, epara, epos)

# Python COM: 튜플 반환
slist, spara, spos, elist, epara, epos = hwp.GetSelectedPos()
```

> 같은 리스트 내에서만 선택 가능. 본문 ↔ 표셀 간 교차 선택 불가.

---

## SetTextFile (텍스트/HTML 삽입)

```python
hwp.SetTextFile(data, format, option)
# format: "TEXT", "HTML", "HWPML2X", "HWP" (BASE64)
# option: "" = 문서 대체, "insertfile" = 커서 위치에 삽입
```

---

## InsertCtrl (컨트롤 삽입)

```python
# 5x5 표 삽입
act = hwp.CreateAction("TableCreate")
set = act.CreateSet()
act.GetDefault(set)
set.SetItem("Rows", 5)
set.SetItem("Cols", 5)
ctrl = hwp.InsertCtrl("tbl", set)
```

주요 ctrlid: `"tbl"` (표), `"gso"` (그리기), `"fn "` (각주), `"en "` (미주), `"eqed"` (수식)

---

## KeyIndicator (커서 위치 정보)

```python
seccnt, secno, pageno, colno, line, pos, over, ctrlname = hwp.KeyIndicator()
# ctrlname: 커서 위치의 컨트롤 이름 (예: "표", "그림")
```

---

## RunScriptMacro (매크로 실행)

```python
hwp.RunScriptMacro("OnScriptMacro_함수명", 1, 0)
# 1 = JavaScript, 0 = 실행
```

> 매크로 보안 설정에 따라 차단될 수 있음. 데스크톱 COM 전용.

---

## 데스크톱 COM 전용 메서드

다음 메서드들은 WebHwp에서 사용 불가:
- FindPrivateInfo / ProtectPrivateInfo
- XHwpDocuments / XHwpWindows
- KeyIndicator
- GetMousePos
- SetBarCodeImage
- RunScriptMacro

---

## 출처

- https://developer.hancom.com/webhwp/devguide/hwpctrl/methods
- https://wikidocs.net/book/8956
- https://pyhwpx.com/
- https://forum.developer.hancom.com/
