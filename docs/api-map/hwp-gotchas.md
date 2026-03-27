# HWP COM API — LLM이 모르는 것들

> 공식 문서나 학습 데이터만으로는 알 수 없는, 실험으로 밝혀낸 사실들.
> 한글 2018 (10.0.0.5060), HwpObject (OLE Automation) 기준.

---

## 1. pos는 글자 인덱스가 아니다

SetPos, MovePos, SelectText의 `pos` 파라미터는 **HWP 내부 오프셋**이다.

- 첫 문단은 섹션/칼럼 컨트롤 때문에 오프셋 존재 (기본 문서에서 16)
- 이후 문단은 0부터
- 영문/한글/특수문자 모두 **글자당 1**
- 오프셋은 문서 구조에 따라 변동 가능 → **하드코딩 금지, 동적 감지 필수**

```js
// ❌ LLM이 흔히 쓰는 코드 — 안 됨
hwp.SetPos(0, 0, 5);           // "5번째 글자"가 아님
hwp.SelectText(0, 2, 0, 5);   // 가짜 선택 → Delete 안 먹힘

// ✅ 올바른 코드 — 동적으로 GetPosBySet().Item("Pos")로 읽어야 함
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

---

## 2. VT_VARIANT 언래핑

`GetPosBySet().Item("Pos")` 등이 `VT_VARIANT`로 감싸져 반환됨. 브릿지에서 `VT_VARIANT` → 재귀 언래핑을 처리하지 않으면 null이 됨.

**이 프로젝트에서 수정 완료:**
- `VT_VARIANT` 재귀 언래핑
- `VT_UI4/VT_UI2/VT_I1/VT_UI1` 추가
- `VT_BYREF` 포인터 역참조

---

## 3. BreakPara가 자동교정을 발동한다

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

## 4. SelectText도 pos 오프셋이 필요하다

`SelectText(spara, spos, epara, epos)`의 spos/epos도 SetPos와 동일한 **내부 오프셋**을 사용한다.
글자 인덱스를 넣으면 텍스트가 아닌 영역을 선택하게 되어 saveblock/Delete가 빈 결과를 반환한다.

```js
// ❌ 글자 인덱스 사용 → 빈 선택
hwp.SelectText(0, 0, 0, 5);
hwp.GetTextFile("UNICODE", "saveblock");  // "" (빈 문자열)

// ✅ pos 오프셋 적용 → 정상 선택
var offset = getParaOffset(0);  // 보통 16
hwp.SelectText(0, offset, 0, offset + 5);
hwp.GetTextFile("UNICODE", "saveblock");  // "ABCDE" ✅
hwp.GetTextFile("HWPML2X", "saveblock");  // XML로도 가능 ✅
hwp.Run("Delete");  // 삭제도 정상 ✅
```

**이전에 "SelectText로 선택하면 saveblock이 안 된다"고 판단했던 것은 오진.
실제 원인은 pos 인코딩을 몰랐기 때문.**

---

## 5. SetTextFile은 교체가 아니라 삽입

```js
// ❌ 공식 문서: ""은 전체 교체 — 실제로는 삽입
hwp.SetTextFile("새 내용", "UNICODE", "");  // 기존 내용 뒤에 추가됨

// ✅ 전체 교체 패턴
hwp.Run("SelectAll");
hwp.Run("Delete");
hwp.SetTextFile("새 내용", "UNICODE", "");
```

option=""이든 "insertfile"이든 동일하게 커서 위치에 삽입. 교체하려면 SelectAll → Delete → SetTextFile.

---

## 6. GetTextFile("UTF8") 미지원

null 반환. **UNICODE만 사용.**

| Format | 결과 |
|--------|------|
| `"UNICODE"` | ✅ 유일하게 안전 |
| `"TEXT"` | CP949 — Node.js에서 한글 깨짐 |
| `"UTF8"` | null 반환 (미지원) |
| `"HTML"` | ✅ 서식 포함 |
| `"HWPML2X"` | ✅ 완전한 XML |
| `"HWP"` | ✅ Base64 바이너리 |

---

## 7. WinBrushFaceStyle = -1

셀 배경색 설정 시 WinBrushFaceStyle을 **-1**로 해야 단색. 0 이상은 빗금 패턴.

```js
// ❌ FaceStyle=0 또는 1 → 빗금 패턴이 됨
fa.WinBrushFaceStyle = 0;  // 가로줄 빗금
fa.WinBrushFaceStyle = 1;  // 세로줄 빗금

// ✅ 단색 채우기
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = hwp.RGBColor(255, 0, 0);
fa.WinBrushFaceStyle = -1;  // ← 반드시 -1
```

Windows GDI HBRUSH 규격: -1=단색, 0=가로줄, 1=세로줄, 2~5=대각선/격자.

---

## 8. 복잡한 표 진입 실패

MoveDown으로 단순 표는 진입 가능하지만, **이력서 같은 복잡한 표는 실패**.

```js
// ❌ 복잡한 표에서 MoveDown 진입 실패
hwp.Run("MoveDocBegin");
hwp.Run("MoveDown");  // 표를 건너뛰거나 엉뚱한 셀로 감

// ✅ HWPML2X 라운드트립이 유일한 안정적 방법
var xml = hwp.GetTextFile("HWPML2X", "");
// XML 파싱 → 셀 내용 수정 → SetTextFile로 복원
```

---

## 9. InsertPicture 크기

**항상 원본 크기로 삽입.** Width/Height 파라미터 무시됨.

```js
// 삽입 후 별도 크기 조절 필요
var ctrl = hwp.InsertPicture(path, 1, 0, 0, 0, 0, 0, 0);
// ctrl Properties로 Width/Height 설정
```

- 8개 파라미터 전부 필수 (부족하면 에러)
- 로컬 파일만 가능 (URL 안 됨)
- 반환값은 Dispatch 객체 (문자열 변환 불가)

---

## 10. SaveAs 보안모듈 필요

Open은 보안모듈 없이 가능 (blocking popup — "모두 허용" 후 세션 동안 팝업 안 뜸).
**SaveAs는 보안모듈 없이 무조건 실패** (non-blocking error 0x80010105).

---

## 11. AllReplace 빈 문자열

빈 문자열로 치환 시 실패할 수 있음 (0x80010105 에러).

```js
// ❌ 빈 문자열 → 에러
hwp.HParameterSet.HFindReplace.ReplaceString = "";

// ✅ 공백으로 치환
hwp.HParameterSet.HFindReplace.ReplaceString = " ";
```

---

## 12. List ID

표의 각 셀은 고유 **List ID**를 가짐. `SetPos(list, 0, 0)`으로 직접 점프 가능.

**하지만:**
- List ID는 **런타임에서만** 알 수 있고 예측 불가능
- HWPML2X/HTML에 List ID 정보 없음
- 세션마다 달라질 수 있음 → 매번 스캔 필요

---

## 13. GetDefault는 현재값이 아닌 기본값

`HAction.GetDefault`는 **기본값으로 리셋**한다. 현재 서식을 읽는 용도가 아님.

```js
// ❌ 현재 서식 읽기 — 전부 0 반환
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
hwp.HParameterSet.HCharShape.Height;  // 0 (실제 값 아님)

// ✅ 서식 읽기 — 루트 프로퍼티 사용
hwp.CharShape.Item("Height");    // 2400 (실제 값)
hwp.ParaShape.Item("AlignType"); // 3 (실제 값)
```

CharShape, CellBorderFill 등 모두 동일. **GetDefault 리셋 주의**: LineSpacing만 바꾸면 다른 프로퍼티가 기본값으로 돌아감. 관련 프로퍼티를 **전부 명시적으로 설정**할 것.

---

## 14. 프로퍼티 이름이 공식 문서와 다르다

| 공식/예상 | 실제 | 대상 |
|----------|------|------|
| `Alignment` | **`AlignType`** | HParaShape |
| `SpaceBeforePara` | **`PrevSpacing`** | HParaShape |
| `SpaceAfterPara` | **`NextSpacing`** | HParaShape |
| `Underline` | **`UnderlineType`** | HCharShape |
| `StrikeOut` | **`StrikeOutType`** | HCharShape |
| `BorderColorLeft` | **`BorderCorlorLeft`** | HCellBorderFill (오타 아님, API가 이렇게 씀) |

---

## 15. RGBColor는 BGR 순서

```js
hwp.RGBColor(255, 0, 0);  // = 0x000000FF = 255 (빨강)
hwp.RGBColor(0, 0, 255);  // = 0x00FF0000 = 16711680 (파랑)
```

Windows COLORREF 형식 (BGR). RGB 순서가 아님.

---

## 16. CreatePageImage — BMP/GIF만 작동

```js
hwp.CreatePageImage(path, 0, 150, 24, "BMP");  // ✅
hwp.CreatePageImage(path, 0, 150, 24, "GIF");  // ✅
hwp.CreatePageImage(path, 0, 300, 24, "PNG");  // ❌ ret=true인데 파일 없음
hwp.CreatePageImage(path, 0, 300, 24, "JPG");  // ❌ 동일
```

---

## 17. 필드는 첫 번째만 접근 가능

CreateField로 여러 개 만들 수 있고, GetFieldList/FieldExist에서 전부 보이지만,
**PutFieldText/GetFieldText/MoveToField는 첫 번째 필드만 작동**. 한글 2018 제한.

템플릿 채우기에는 AllReplace 패턴(`{{name}}` → `홍길동`)이 더 실용적.
