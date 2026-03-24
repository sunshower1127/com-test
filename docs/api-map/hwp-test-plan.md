# HWP API 호출 테스트 계획

> 원칙: 각 단계마다 코드를 실행하고, 사용자가 눈으로 확인한 결과만 기록한다.

## 테스트 방법

각 단계별로 Rust 예제(또는 com-cli)로 코드를 실행하고, 사용자가 결과를 확인.
결과는 `hwp-test-results.md`에 ✅/❌/⚠️로 기록.

---

## 단계 0: 기본 연결

**의존**: 없음 (모든 테스트의 전제)

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 0-1 | `DispatchObject::create("HWPFrame.HwpObject")` | 에러 없이 반환 |
| 0-2 | `XHwpWindows.Item(0).Visible = true` | 한글 창이 화면에 보이는가 |
| 0-3 | `Version` 프로퍼티 읽기 | 버전 문자열 출력 |
| 0-4 | `EditMode` 프로퍼티 읽기 | 정수값 반환 |
| 0-5 | `IsEmpty` 프로퍼티 읽기 | true/false |
| 0-6 | `PageCount` 프로퍼티 읽기 | 정수값 반환 |
| 0-7 | `SetMessageBoxMode(1)` | 에러 없이 실행 (팝업 억제) |

---

## 단계 1: 문서 생명주기

**의존**: 단계 0 성공

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 1-1 | `Run("FileNew")` | 새 빈 문서가 열리는가 |
| 1-2 | `Open("test.hwp")` — 기존 파일 | 문서가 열리고 내용이 보이는가 |
| 1-3 | `Path` 프로퍼티 읽기 | 열린 파일 경로가 반환되는가 |
| 1-4 | `IsModified` 읽기 | false (방금 열었으므로) |
| 1-5 | `SaveAs("output.hwp", "HWP")` | 파일이 생성되는가 |
| 1-6 | `SaveAs("output.hwpx", "HWPX")` | HWPX 형식 저장 |
| 1-7 | `SaveAs("output.pdf", "PDF")` | PDF 형식 저장 |
| 1-8 | `SaveAs("output.txt", "TEXT")` | 텍스트 형식 저장 |
| 1-9 | `Save(false)` | 현재 파일에 저장 |
| 1-10 | `Clear(1)` | 문서 닫힘 (저장 안 함) |
| 1-11 | `XHwpDocuments.Add(false)` | 새 탭/문서 추가 |
| 1-12 | `XHwpDocuments.Count` | 문서 수 반환 |
| 1-13 | `Quit()` | 한글 종료 |

---

## 단계 2: 텍스트 입출력 (기본)

**의존**: 단계 1 (문서 열기/닫기 가능)

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 2-1 | CreateAction("InsertText") + SetItem("Text", "안녕하세요") + Execute | 텍스트가 화면에 보이는가 |
| 2-2 | 여러 줄 삽입: "첫째줄\r\n둘째줄" | 줄바꿈이 되는가 |
| 2-3 | `GetTextFile("UNICODE", "")` | 삽입한 텍스트가 그대로 반환되는가 |
| 2-4 | `GetTextFile("UTF8", "")` | UTF8로도 동작하는가 |
| 2-5 | `SetTextFile("새 내용", "UNICODE", "")` | 문서 전체가 교체되는가 |
| 2-6 | `GetPageText(0, "")` | 첫 페이지 텍스트 반환 |
| 2-7 | `GetHeadingString` | 제목 문자열 반환 |
| 2-8 | `Run("BreakPara")` — 줄바꿈 | 새 문단이 생기는가 |

---

## 단계 3: 커서 이동 + 선택

**의존**: 단계 2 (텍스트가 있는 상태)

미리 여러 줄 텍스트를 넣어두고 테스트.

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 3-1 | `Run("MoveDocBegin")` | 커서가 문서 처음으로 |
| 3-2 | `Run("MoveDocEnd")` | 커서가 문서 끝으로 |
| 3-3 | `Run("MoveLineBegin")` | 현재 줄 처음 |
| 3-4 | `Run("MoveLineEnd")` | 현재 줄 끝 |
| 3-5 | `Run("MoveNextPara")` | 다음 문단 |
| 3-6 | `Run("MovePrevPara")` | 이전 문단 |
| 3-7 | `Run("MoveSelDocBegin")` | 문서 처음까지 선택 |
| 3-8 | `Run("MoveSelDocEnd")` | 문서 끝까지 선택 |
| 3-9 | `Run("MoveSelLineEnd")` | 줄 끝까지 선택 |
| 3-10 | `Run("MoveSelLineBegin")` | 줄 처음까지 선택 |
| 3-11 | `SelectText(0, 0, 0, 5)` | 첫 문단의 0~5 위치 선택되는가 |
| 3-12 | `MovePos(2, 0, 0)` | 특정 위치로 이동 (moveID 값 확인 필요) |
| 3-13 | `SetPos(0, 0, 0)` + `GetPosBySet` | 위치 설정 후 읽기 |
| 3-14 | `GetSelectedPosBySet` | 선택 영역 위치 읽기 |

---

## 단계 4: 텍스트 수정 (핵심!)

**의존**: 단계 3 (선택이 가능한 상태)

> 이 단계가 "문서 수정"이 되는지 안 되는지의 핵심.

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 4-1 | 텍스트 선택 → `Run("Delete")` | 선택한 텍스트가 삭제되는가 |
| 4-2 | 텍스트 선택 → InsertText로 덮어쓰기 | 선택 영역이 새 텍스트로 교체되는가 |
| 4-3 | `SelectText(0,0,0,5)` → Delete → InsertText | 프로그래밍으로 부분 교체 가능한가 |
| 4-4 | AllReplace: FindString="A", ReplaceString="B" | 찾기/바꾸기가 작동하는가 |
| 4-5 | AllReplace + IgnoreMessage=1 | 확인 대화상자 없이 실행되는가 |
| 4-6 | SetTextFile로 전체 교체 후 원래 내용 비교 | 전체 덮어쓰기 패턴이 안정적인가 |
| 4-7 | `Run("Undo")` | 되돌리기 작동하는가 |
| 4-8 | `Run("Redo")` | 다시 실행 작동하는가 |

---

## 단계 5: 글자 서식 (CharShape)

**의존**: 단계 3 (선택 가능) + 단계 2 (텍스트 입력 가능)

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 5-1 | 텍스트 선택 → CharShape: Height=2400 (24pt) | 글자 크기가 커지는가 |
| 5-2 | CharShape: Bold=1 | 볼드 적용되는가 |
| 5-3 | CharShape: Italic=1 | 이탤릭 적용되는가 |
| 5-4 | CharShape: TextColor=RGBColor(255,0,0) | 빨간색으로 변하는가 |
| 5-5 | CharShape: Underline=1 | 밑줄 적용되는가 |
| 5-6 | CharShape: StrikeOut=1 | 취소선 적용되는가 |
| 5-7 | 빈 커서에서 CharShape 설정 → InsertText | 이후 입력에 서식이 적용되는가 |
| 5-8 | `hwp.CharShape` 프로퍼티로 현재 서식 읽기 | Item("Height") 등으로 값이 나오는가 |

---

## 단계 6: 문단 서식 (ParagraphShape)

**의존**: 단계 2 + 단계 3

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 6-1 | ParagraphShape: Alignment=3 (가운데) | 문단 정렬 변경되는가 |
| 6-2 | ParagraphShape: Alignment=2 (오른쪽) | 오른쪽 정렬 |
| 6-3 | ParagraphShape: LineSpacing 변경 | 줄간격이 변하는가 |
| 6-4 | ParagraphShape: LeftMargin 변경 | 왼쪽 들여쓰기 |
| 6-5 | ParagraphShape: SpaceBeforePara / SpaceAfterPara | 문단 간격 |
| 6-6 | `hwp.ParaShape` 프로퍼티로 현재 문단 서식 읽기 | 값이 나오는가 |

---

## 단계 7: 표 (Table)

**의존**: 단계 2

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 7-1 | TableCreate: Rows=3, Cols=2, WidthType=0 | 3x2 표가 생기는가 |
| 7-2 | 표 안에서 InsertText | 셀에 텍스트가 들어가는가 |
| 7-3 | `Run("TableRightCell")` | 다음 셀로 이동하는가 |
| 7-4 | `Run("TableLeftCell")` | 이전 셀 |
| 7-5 | `Run("TableUpperCell")` | 위 셀 |
| 7-6 | `Run("TableLowerCell")` | 아래 셀 |
| 7-7 | 표 밖으로 나가기: `Run("MoveDocEnd")` 등 | 표 밖으로 나갈 수 있는가 |
| 7-8 | `hwp.CellShape` 프로퍼티 | 셀 안에서 읽히는가 |
| 7-9 | 표 셀 안에서 CharShape 적용 | 셀 내 서식이 되는가 |

---

## 단계 8: 그림 삽입

**의존**: 단계 1

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 8-1 | `InsertPicture(path, true, 0, false, false, 0, 0, 0)` | 그림이 삽입되는가 |
| 8-2 | `InsertBackgroundPicture(...)` | 배경 그림 삽입 |
| 8-3 | `CreatePageImage(path, 0, 300, 24, "PNG")` | 페이지를 이미지로 저장 |

---

## 단계 9: 필드

**의존**: 단계 2

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 9-1 | `CreateField("", "", "myfield")` | 필드 생성 |
| 9-2 | `FieldExist("myfield")` | true 반환 |
| 9-3 | `PutFieldText("myfield", "필드내용")` | 필드에 텍스트 설정 |
| 9-4 | `GetFieldText("myfield")` | "필드내용" 반환 |
| 9-5 | `GetFieldList(0, 0)` | 필드 목록 반환 |
| 9-6 | `MoveToField("myfield", true, true, false)` | 필드로 이동 |

---

## 단계 10: 페이지 설정

**의존**: 단계 1

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 10-1 | PageSetup → 여백 변경 (SetItem 키 확인 필요) | 여백이 바뀌는가 |
| 10-2 | PageSetup → 용지 방향 변경 | 가로/세로 전환되는가 |
| 10-3 | PageSetup → 용지 크기 변경 | A4 → B5 등 |

---

## 단계 11: 컨트롤 순회 (HeadCtrl/LastCtrl)

**의존**: 단계 7 또는 단계 8 (컨트롤이 있는 문서)

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 11-1 | `HeadCtrl` → CtrlID, UserDesc 읽기 | 첫 컨트롤 정보 |
| 11-2 | `HeadCtrl.Next` 반복 | 모든 컨트롤 순회 가능한가 |
| 11-3 | `LastCtrl` 읽기 | 마지막 컨트롤 |
| 11-4 | 컨트롤의 `Properties` 읽기 | ParameterSet 반환 |
| 11-5 | `DeleteCtrl(ctrl)` | 컨트롤 삭제 가능한가 |

---

## 단계 12: HParameterSet 직접 접근 vs CreateAction

**의존**: 단계 2

> 두 가지 패턴을 동일 작업으로 비교하여 어느 게 되고 안 되는지 확인.

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 12-1 | `hwp.HParameterSet` → `HInsertText` 접근 | Dispatch가 나오는가 / 에러인가 |
| 12-2 | HParameterSet 방식으로 InsertText 시도 | 텍스트가 삽입되는가 |
| 12-3 | `hwp.HAction.GetDefault("InsertText", set)` | HAction 경유 방식 |
| 12-4 | CreateAction 방식으로 동일 작업 | 비교 |

---

## 단계 13: 유틸리티

**의존**: 없음

| # | 테스트 | 확인 방법 |
|---|--------|-----------|
| 13-1 | `RGBColor(255, 0, 0)` | 정수값 반환 |
| 13-2 | `MiliToHwpUnit(25.4)` | 변환값 반환 |
| 13-3 | `PointToHwpUnit(12.0)` | 변환값 반환 |
| 13-4 | `HwpLineType("Solid")` 등 열거형 변환 | 정수값 반환 |

---

## 실행 순서 요약

```
단계 0  기본 연결          ← 모든 것의 전제
  ↓
단계 1  문서 생명주기       ← Open/Save/Close
  ↓
단계 2  텍스트 입출력       ← Insert/Get
  ↓
단계 3  커서 이동+선택      ← Move/Select
  ↓
단계 4  텍스트 수정 ★      ← Delete/Replace (핵심!)
  ↓
단계 5  글자 서식           ← CharShape
  ↓
단계 6  문단 서식           ← ParagraphShape
  ↓
단계 7  표                 ← TableCreate
  ↓
단계 8  그림 삽입           ← InsertPicture
  ↓
단계 9  필드               ← Field 시스템
  ↓
단계 10 페이지 설정         ← PageSetup
  ↓
단계 11 컨트롤 순회         ← HeadCtrl/Next
  ↓
단계 12 HParameterSet 비교  ← 두 패턴 비교
  ↓
단계 13 유틸리티            ← 독립 실행 가능
```
