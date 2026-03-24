# HWP COM API — LLM이 모르는 것들

> 공식 문서나 학습 데이터만으로는 알 수 없는, 실험으로 밝혀낸 사실들.
> 한글 2018 (10.0.0.5060), HwpObject (OLE Automation) 기준.

---

## 1. pos는 글자 인덱스가 아니다

SetPos, MovePos, SelectText의 `pos` 파라미터는 **HWP 내부 오프셋**이다.

```js
// ❌ LLM이 흔히 쓰는 코드 — 안 됨
hwp.SetPos(0, 0, 5);           // "5번째 글자"가 아님
hwp.SelectText(0, 2, 0, 5);   // 가짜 선택 → Delete 안 먹힘

// ✅ 올바른 코드
function getParaOffset(para) {
  hwp.SetPos(0, para, 0);
  hwp.Run("MoveParaBegin");
  var ps = hwp.GetPosBySet();
  return ps.Item("Pos");  // 첫 문단: 16(보통), 이후 문단: 0
}
var offset = getParaOffset(0);
hwp.SetPos(0, 0, offset + 5);
hwp.SelectText(0, offset + 2, 0, offset + 5);  // 정상 선택
```

- 첫 문단은 섹션/칼럼 컨트롤 때문에 오프셋 존재 (기본 문서에서 16)
- 이후 문단은 0부터
- 영문/한글/특수문자 모두 **글자당 1**
- 오프셋은 문서 구조에 따라 변동 가능 → **하드코딩 금지, 동적 감지 필수**

---

## 2. GetDefault는 현재 값을 안 읽는다

`HAction.GetDefault`는 **기본값으로 리셋**한다. 현재 서식을 읽는 용도가 아님.

```js
// ❌ 현재 서식 읽기 — 전부 0 반환
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Height;  // 0 (실제 값 아님)

// ✅ 서식 읽기 — 루트 프로퍼티 사용
hwp.CharShape.Item("Height");    // 2400 (실제 값)
hwp.ParaShape.Item("AlignType"); // 3 (실제 값)
```

**GetDefault 리셋 주의**: LineSpacing만 바꾸면 다른 프로퍼티가 기본값으로 돌아감. 관련 프로퍼티를 **전부 명시적으로 설정**할 것.

```js
// ❌ LineSpacing만 설정 → 다른 값 리셋
hwp.HAction.GetDefault("ParagraphShape", ...);
hwp.HParameterSet.HParaShape.LineSpacing = 250;  // 읽으면 160

// ✅ 관련 프로퍼티 함께 설정
hwp.HParameterSet.HParaShape.LineSpacingType = 0;
hwp.HParameterSet.HParaShape.LineSpacing = 250;  // 정상
```

---

## 3. 프로퍼티 이름이 공식 문서와 다르다

| 공식/예상 | 실제 | 대상 |
|----------|------|------|
| `Alignment` | **`AlignType`** | HParaShape |
| `SpaceBeforePara` | **`PrevSpacing`** | HParaShape |
| `SpaceAfterPara` | **`NextSpacing`** | HParaShape |
| `Underline` | **`UnderlineType`** | HCharShape |
| `StrikeOut` | **`StrikeOutType`** | HCharShape |
| `BorderColorLeft` | **`BorderCorlorLeft`** | HCellBorderFill (오타 아님, API가 이렇게 씀) |

---

## 4. BreakPara가 자동교정을 발동한다

`Run("BreakPara")` 실행 시 커서 문단에 HWP 빠른 교정(Quick Correct)이 자동 발동.
예: `"첫번째"` → `"첫 번째"`

- API로 끄는 방법 **없음** (AutoSpellCheck, ToggleAutoCorrect 모두 효과 없음)
- **우회: `\r\n`을 InsertText에 포함** → 동일한 문단 구조, 교정 안 됨

```js
// ❌ 자동교정 발동
insert("첫번째");
hwp.Run("BreakPara");

// ✅ 자동교정 안 됨
insert("첫번째\r\n두번째");
```

**단, 문서 중간에서 줄 나누기는 BreakPara 필수** (insert `\r\n`은 중간 삽입 시 다르게 동작).

---

## 5. SetTextFile option=""은 교체가 아니라 삽입

```js
// ❌ 공식 문서: ""은 전체 교체 — 실제로는 삽입
hwp.SetTextFile("새 내용", "UNICODE", "");  // 기존 내용 뒤에 추가됨

// ✅ 전체 교체 패턴
hwp.Run("SelectAll");
hwp.Run("Delete");
hwp.SetTextFile("새 내용", "UNICODE", "");
```

---

## 6. 셀 배경색 — WinBrushFaceStyle 함정

```js
// ❌ FaceStyle=0 또는 1 → 빗금 패턴이 됨
fa.WinBrushFaceStyle = 0;  // 가로줄 빗금
fa.WinBrushFaceStyle = 1;  // 세로줄 빗금

// ✅ 단색 채우기
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = hwp.RGBColor(255, 0, 0);
fa.WinBrushFaceStyle = -1;  // ← 반드시 -1 또는 생략
```

WinBrushFaceStyle은 Windows HBRUSH 패턴: -1=단색, 0=가로줄, 1=세로줄, 2~5=대각선/격자.

---

## 7. InsertPicture는 Dispatch를 반환한다

```js
// ❌ 반환값을 문자열로 쓰면 에러
console.log("결과=" + hwp.InsertPicture(...));  // Cannot convert object to primitive value

// ✅ Dispatch 객체로 받기
var ctrl = hwp.InsertPicture(path, 1, 0, 0, 0, 0, 0, 0);
ctrl.CtrlID;    // "gso"
ctrl.UserDesc;  // "그림"
```

- **8개 파라미터 전부 필수** (부족하면 에러)
- **로컬 파일만 가능** (URL 안 됨)
- 외부 이미지 → 임시파일 저장 후 InsertPicture

---

## 8. 표 밖에서 안으로 못 들어간다 (MoveRight로는)

```js
// ❌ MoveRight는 표를 건너뜀
hwp.Run("MoveDocBegin");
hwp.Run("MoveRight");  // 표 안 들어감

// ✅ MoveDown으로 진입
hwp.Run("MoveDocBegin");
hwp.Run("MoveDown");         // 표 안으로 진입 (아무 셀)
hwp.Run("TableUpperCell");   // 첫 행으로
hwp.Run("TableLeftCell");    // 첫 열로
```

또는 SetPos(list, 0, 0)으로 직접 진입 (list ID를 알 경우).

---

## 9. SaveAs는 보안모듈 없이 안 된다

Open은 팝업 허용으로 가능하지만, **SaveAs는 보안모듈 없이 무조건 실패** (0x80010105).
Open은 블로킹(팝업 대기), SaveAs는 논블로킹(즉시 에러).

---

## 10. GetTextFile — UNICODE만 쓸 것

| Format | 결과 |
|--------|------|
| `"UNICODE"` | ✅ 유일하게 안전 |
| `"TEXT"` | CP949 — Node.js에서 한글 깨짐 |
| `"UTF8"` | null 반환 (미지원) |
| `"HTML"` | ✅ 서식 포함 |
| `"HWPML2X"` | ✅ 완전한 XML |
| `"HWP"` | ✅ Base64 바이너리 |

---

## 11. AllReplace 빈 문자열로 치환 안 됨

```js
// ❌ 빈 문자열 → 0x80010105 에러
hwp.HParameterSet.HFindReplace.ReplaceString = "";

// ✅ 공백으로 치환
hwp.HParameterSet.HFindReplace.ReplaceString = " ";
```

---

## 12. RGBColor는 BGR 순서

```js
hwp.RGBColor(255, 0, 0);  // = 0x000000FF = 255 (빨강)
hwp.RGBColor(0, 0, 255);  // = 0x00FF0000 = 16711680 (파랑)
```

Windows COLORREF 형식 (BGR). RGB 순서가 아님.

---

## 13. CreatePageImage — BMP/GIF만 작동

```js
hwp.CreatePageImage(path, 0, 150, 24, "BMP");  // ✅
hwp.CreatePageImage(path, 0, 150, 24, "GIF");  // ✅
hwp.CreatePageImage(path, 0, 300, 24, "PNG");  // ❌ ret=true인데 파일 없음
hwp.CreatePageImage(path, 0, 300, 24, "JPG");  // ❌ 동일
```

---

## 14. 필드는 첫 번째만 접근 가능

CreateField로 여러 개 만들 수 있고, GetFieldList/FieldExist에서 전부 보이지만,
**PutFieldText/GetFieldText/MoveToField는 첫 번째 필드만 작동**. 한글 2018 제한.

템플릿 채우기에는 AllReplace 패턴(`{{name}}` → `홍길동`)이 더 실용적.

---

## 15. VT_VARIANT — 브릿지 주의사항

HWP ParameterSet의 `Item()` 메서드는 `VT_VARIANT`(중첩 Variant) 안에 값을 감싸서 반환.
브릿지에서 `VT_VARIANT` → 재귀 언래핑을 처리하지 않으면 null이 됨.

이 프로젝트에서 수정 완료:
- `VT_VARIANT` 재귀 언래핑
- `VT_UI4/VT_UI2/VT_I1/VT_UI1` 추가
- `VT_BYREF` 포인터 역참조
