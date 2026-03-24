# HWP COM API 이슈 총정리 (단계 0~4)

> 한글 2018 (10.0.0.5060), 보안모듈 미설치, HwpObject (OLE Automation)

---

## 보안 관련

| 이슈 | 설명 | 우회 |
|------|------|------|
| **SaveAs 보안모듈 필수** | 모든 포맷에서 0x80010105 에러. 논블로킹으로 즉시 실패 | 보안모듈 DLL 설치 + RegisterModule (step-01 TODO 참고) |
| **Save 실질 저장 안 됨** | 에러는 안 나지만 IsModified 여전히 true | 보안모듈 설치 후 재테스트 필요 |
| **Open 보안 팝업** | 블로킹 팝업 — "모두 허용" 후 세션 동안 팝업 안 뜸 | 보안모듈 설치로 해결 가능 |
| **AllReplace 빈 문자열** | ReplaceString="" 시 0x80010105 에러 | 공백(" ")으로 치환하거나 Find→선택→Delete |

---

## 해결된 이슈 (브릿지 수정)

### VT_VARIANT 미처리 → 해결

`GetPosBySet().Item("Pos")` 등이 null을 반환하던 문제. HWP ParameterSet의 `Item()` 메서드가 `VT_VARIANT`(중첩 Variant) 안에 값을 감싸서 반환하는데, 브릿지의 `variant_from_raw`가 이 타입을 처리하지 못했음.

**수정:** `VT_VARIANT` → 재귀 언래핑 + `VT_UI4/VT_UI2/VT_I1/VT_UI1` + `VT_BYREF` 처리 추가.

**영향 범위:** GetPosBySet Item 읽기, ParameterSet Item 전반.

---

## pos 인코딩 — SetPos, MovePos, SelectText 공통

pos 파라미터는 **글자 인덱스가 아니라 HWP 내부 오프셋**을 사용한다. 첫 문단에는 섹션 속성/칼럼 속성 등 컨트롤 문자가 텍스트 앞에 존재하여 오프셋이 발생한다.

| 문단 | pos 공식 | 이유 |
|------|---------|------|
| **첫 문단 (para=0)** | `기본 오프셋 + 글자인덱스` | 섹션/칼럼 컨트롤 문자 (기본 문서에서 16) |
| **이후 문단 (para>0)** | `글자인덱스` | 컨트롤 없음 |

- 영문/한글/특수문자 모두 **글자당 1**
- 오프셋은 문서 구조에 따라 변동 가능 (머리말/꼬리말 등 추가 시)

**안전한 사용법 — 오프셋 동적 감지:**

```js
function getParaOffset(para) {
  hwp.SetPos(0, para, 0);
  hwp.Run("MoveParaBegin");
  var ps = hwp.GetPosBySet();
  return ps.Item("Pos");
}

// SetPos
var offset = getParaOffset(0);
hwp.SetPos(0, 0, offset + 5);  // 첫 문단 6번째 글자

// SelectText
hwp.SelectText(0, offset+2, 0, offset+5);  // "CDE" 선택
hwp.Run("Delete");  // ✅ 정상

// 이후 문단 (오프셋 0)
hwp.SetPos(0, 1, 3);  // 두번째 문단 4번째 글자
hwp.SelectText(1, 1, 1, 3);  // 두번째 문단 "EF" 선택
```

**HWPX 분석으로 확인된 원인:**

```xml
<!-- 첫 문단: 컨트롤이 텍스트 앞에 존재 -->
<hp:p>
  <hp:run><hp:secPr .../><hp:ctrl><hp:colPr .../></hp:ctrl></hp:run>
  <hp:run><hp:t>ABC</hp:t></hp:run>
</hp:p>

<!-- 이후 문단: 바로 텍스트 -->
<hp:p>
  <hp:run><hp:t>DEF</hp:t></hp:run>
</hp:p>
```

---

## 공식 문서와 다른 동작

| 이슈 | 공식 문서 | 실제 동작 | 단계 |
|------|-----------|-----------|:----:|
| **SetTextFile option=""** | 전체 교체 | **커서 위치에 삽입** (교체 아님). HwpCtrl/HwpObject 차이 | 2 |
| **GetTextFile "UTF8"** | 일부 문서에서 언급 | **null 반환** (미지원) | 2 |
| **pos 파라미터** | 글자 인덱스 | HWP 내부 오프셋 (위 섹션 참고) | 3 |

---

## 자동교정 (빠른 교정)

`Run("BreakPara")` 실행 시 **커서가 있는 문단**에 HWP 빠른 교정(Quick Correct)이 발동.
예: `"첫번째"` → `"첫 번째"` (띄어쓰기 자동 삽입)

| 항목 | 내용 |
|------|------|
| **트리거** | `BreakPara`만. 다른 커서 이동/선택은 발동 안 함 |
| **범위** | 커서 있는 문단만 |
| **API로 끄기** | **불가** — AutoSpellCheck, ToggleAutoCorrect, HQCorrect, 레지스트리 모두 효과 없음 |
| **우회** | `InsertText`에 `\r\n` 포함 → BreakPara와 동일한 문단 구조, 교정 안 됨 |

```js
// ❌ 자동교정 발동
insert('첫번째');
hwp.Run('BreakPara');

// ✅ 자동교정 안 됨, 동일한 문단 구조
insert('첫번째\r\n두번째');
```

**주의:** `insert('\r\n')`은 문서 끝에서만 정상. 문서 중간에서는 BreakPara 필요 (자동교정 감수).

---

## 인코딩/포맷 관련

| 이슈 | 설명 | 단계 |
|------|------|:----:|
| **GetTextFile "TEXT"** | CP949/ANSI로 반환 → Node.js에서 한글 깨짐 | 2 |
| **GetTextFile "UTF8"** | null 반환 (미지원) | 2 |
| **RGBColor** | BGR 순서 (Windows COLORREF). RGBColor(255,0,0) = 0x000000FF | 0 |

**텍스트 추출은 `"UNICODE"` 만 사용.**

---

## 작동하지 않는 API 목록

| API | 증상 | 대체 |
|-----|------|------|
| `SaveAs` / `Save` | 보안모듈 없이 에러 | 보안모듈 설치 |
| `GetTextFile("UTF8")` | null | `"UNICODE"` 사용 |
| `AllReplace` 빈 문자열 | 0x80010105 | 공백 치환 |
| `BackSpace` (문서 끝) | 안 먹힘 | `Delete` 사용 |

---

## 정상 작동 확인된 핵심 패턴

```js
// 텍스트 삽입 (\r\n으로 줄바꿈 — 자동교정 회피)
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
hwp.HParameterSet.HInsertText.Text = "텍스트\r\n다음줄";
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);

// 텍스트 추출
hwp.GetTextFile("UNICODE", "");

// 전체 교체
hwp.Run("SelectAll");
hwp.Run("Delete");
hwp.SetTextFile(newContent, "UNICODE", "");

// 부분 교체 — 방법 A: Run 기반
hwp.Run("MoveDocBegin");
for (...) hwp.Run("MoveRight");
for (...) hwp.Run("MoveSelRight");
hwp.InsertText("new");

// 부분 교체 — 방법 B: SelectText + pos 인코딩
var offset = getParaOffset(0);
hwp.SelectText(0, offset+3, 0, offset+6);
hwp.InsertText("new");

// 찾아 바꾸기
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "old";
hwp.HParameterSet.HFindReplace.ReplaceString = "new";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);

// 위치 읽기/이동
var ps = hwp.GetPosBySet();
var list = ps.Item("List");  // 0
var para = ps.Item("Para");  // 문단 번호
var pos = ps.Item("Pos");    // 내부 오프셋

// 위치 저장/복원
var saved = hwp.GetPosBySet();
// ... 다른 작업 ...
hwp.SetPosBySet(saved);

// Undo / Redo
hwp.Run("Undo");
hwp.Run("Redo");
```
