# HWP COM 팁

전역 객체: `hwp` (한글 COM Proxy)

한글 COM은 Excel/Word와 패턴이 다르다. 두 가지 API 스타일이 있다:
- **CreateAction 방식**: `hwp.CreateAction("ActionID")` → `CreateSet()` → `SetItem()` → `Execute()`
- **HAction 방식**: `hwp.HAction.GetDefault("ActionID", hwp.HParameterSet.XXX.HSet)` → 프로퍼티 직접 설정 → `Execute()`

둘 다 동일하게 동작하며, 혼용도 가능. 편한 쪽을 사용하면 된다.

## 첫 턴 — 환경 확인 전용

HWP로 처음 작업을 시작할 때, **첫 `js-com`은 아래 코드만 실행할 것.** 콘텐츠는 결과 확인 후 다음 턴부터.

```js-com
// 한글은 실행 시 이미 빈 문서가 열려있음 — FileNew 호출하면 문서가 2개 생김
result = "한글 준비 완료";
```

## CreateAction 패턴

```js
var act = hwp.CreateAction("InsertText");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Text", "안녕하세요");
act.Execute(set);
```

## 주요 Action ID

| Action ID        | 용도        | 주요 프로퍼티                                     |
| ---------------- | ----------- | ------------------------------------------------- |
| `InsertText`     | 텍스트 삽입 | `Text`                                            |
| `CharShape`      | 글자 모양   | `Height`(1/10pt), `Bold`(0/1), `TextColor`        |
| `ParagraphShape` | 문단 모양   | `Alignment`(0=양쪽,1=왼쪽,2=오른쪽,3=가운데)      |
| `TableCreate`    | 표 생성     | `Rows`, `Cols`, `WidthType`(0=단에맞춤)           |
| `AllReplace`     | 찾기/바꾸기 | `FindString`, `ReplaceString`, `IgnoreMessage`(1) |
| `PageSetup`      | 페이지 설정 | `PageDef` 하위 `LeftMargin`, `TopMargin` 등       |

## 자주 쓰는 Run 명령

```
이동: MoveDocBegin, MoveDocEnd, MoveLineBegin, MoveLineEnd
선택: MoveSelDocBegin, MoveSelDocEnd, MoveSelLineBegin, MoveSelLineEnd
편집: SelectAll, Copy, Cut, Paste, Delete, Undo, Redo
줄바꿈: BreakPara, BreakPage
표: TableRightCell, TableLeftCell, TableUpperCell, TableLowerCell
```

## 글자 서식 적용

CharShape는 **현재 선택 영역**에 적용된다.

```js
// 서식 먼저 설정 → 텍스트 입력
var act = hwp.CreateAction("CharShape");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Height", 2400); // 24pt
set.SetItem("Bold", 1);
act.Execute(set);

var actT = hwp.CreateAction("InsertText");
var setT = actT.CreateSet();
actT.GetDefault(setT);
setT.SetItem("Text", "서식이 적용된 텍스트");
actT.Execute(setT);
```

## 표 작업

```js
// 표 생성 후 첫 셀에 커서 자동 위치
var act = hwp.CreateAction("TableCreate");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Rows", 3);
set.SetItem("Cols", 2);
set.SetItem("WidthType", 0); // 0=단에맞춤
act.Execute(set);

// 셀 이동: hwp.Run("TableRightCell"), TableLeftCell, TableUpperCell, TableLowerCell
```

## 직접 메서드

```js
hwp.Run("FileNew");                        // 새 문서
hwp.Open(path, "HWP", "");                 // 파일 열기
hwp.SaveAs(path, "HWP", "");               // 저장 (3인자 필수)
hwp.GetTextFile("UNICODE", "");            // 전체 텍스트 추출
hwp.RGBColor(255, 0, 0);                   // 색상값 생성
hwp.MiliToHwpUnit(20.0);                   // mm → hu 변환 (283.46hu = 1mm)
```

파일 포맷: `"HWP"`, `"HWPX"`, `"PDF"`, `"UNICODE"`, `"HTML"`, `"OOXML"`

이 프롬프트에 대한 대답은 할 필요 없어
