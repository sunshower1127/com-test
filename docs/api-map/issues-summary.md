# HWP COM API 이슈 총정리

> 한글 2018 (10.0.0.5060), HwpObject (OLE Automation) 기준.
> 테스트를 통해 발견된 모든 이슈와 현재 상태.

---

## ✅ 해결됨

| # | 이슈 | 해결 방법 |
|---|------|-----------|
| 1 | **VT_VARIANT 미처리** — GetPosBySet().Item("Pos") 등이 null 반환 | 브릿지에서 VT_VARIANT 재귀 언래핑 + VT_UI4/VT_UI2/VT_I1/VT_UI1 + VT_BYREF 처리 추가 |
| 2 | **pos 인코딩 이해** — SetPos/MovePos/SelectText의 pos가 글자 인덱스가 아님 | 동적으로 GetPosBySet().Item("Pos")로 오프셋 감지. 첫 문단은 16+글자인덱스, 이후 문단은 0+글자인덱스 |
| 3 | **GetDefault 오해** — 현재 서식 읽기로 오인 | 루트 프로퍼티(hwp.CharShape.Item(), hwp.ParaShape.Item()) 사용으로 전환 |
| 4 | **프로퍼티 이름 불일치** — AlignType, PrevSpacing 등 | 실제 이름 매핑 테이블 작성 완료 (hwp-gotchas.md 참고) |
| 5 | **복잡한 표 데이터 채우기** — COM 셀 이동 불안정 | HWPML2X 라운드트립 파이프라인 확립 (usecase-template-fill.md 참고) |

---

## ⚠️ 우회 가능

| # | 이슈 | 우회 방법 |
|---|------|-----------|
| 1 | **BreakPara 자동교정** — BreakPara가 빠른 교정 트리거 | InsertText에 `\r\n` 포함으로 대체 (문서 중간에서는 BreakPara 필수, 교정 감수) |
| 2 | **SetTextFile은 교체가 아니라 삽입** — option=""도 삽입 동작 | SelectAll → Delete → SetTextFile 패턴 사용 |
| 3 | **GetTextFile("UTF8") 미지원** — null 반환 | "UNICODE" 사용 |
| 4 | **AllReplace 빈 문자열** — 빈 문자열 치환 시 0x80010105 에러 | 공백(" ")으로 치환하거나 Find → 선택 → Delete |
| 5 | ~~**SelectText 선택 후 Delete 안 됨**~~ → **해결!** pos 오프셋 미적용이 원인 | pos 오프셋 적용하면 SelectText도 saveblock/Delete 정상 작동 |
| 6 | **InsertPicture 크기 무시** — 항상 원본 크기로 삽입 | 삽입 후 ctrl Properties로 Width/Height 별도 설정 |
| 7 | **WinBrushFaceStyle 함정** — 0 이상은 빗금 패턴 | WinBrushFaceStyle = -1로 단색 설정 |
| 8 | **Open 보안 팝업** — 블로킹 팝업 표시 | "모두 허용" 클릭 후 세션 동안 재발 안 함. 보안모듈 설치로 완전 해결 가능 |
| 9 | **복잡한 표 진입 실패** — MoveDown으로 이력서 등 복잡한 표 진입 불가 | HWPML2X 라운드트립이 유일한 안정적 방법 |
| 10 | **이미지 삽입 (표 내부)** — HWPML2X 라운드트립으로는 텍스트만 가능 | 플레이스홀더({{PHOTO}}) 삽입 → Find → 삭제 → InsertPicture |
| 11 | **List ID 예측 불가** — 셀 List ID는 런타임에서만 알 수 있음 | 매번 스캔하거나, HWPML2X 라운드트립으로 우회 |

---

## ❌ 미해결/제한사항

| # | 이슈 | 설명 |
|---|------|------|
| 1 | **SaveAs 보안모듈 필수** — 보안모듈 없이 0x80010105 에러 (non-blocking) | 보안모듈 DLL 설치 + RegisterModule 필요 |
| 2 | **Save 실질 저장 안 됨** — 에러 없지만 IsModified 여전히 true | 보안모듈 설치 후 재테스트 필요 |
| 3 | **CreatePageImage PNG/JPG 미작동** — ret=true인데 파일 생성 안 됨 | BMP/GIF만 사용 가능 |
| 4 | **필드 접근 제한** — PutFieldText/GetFieldText가 첫 번째 필드만 작동 | 한글 2018 제한. AllReplace 패턴으로 대체 |
| 5 | **BreakPara 자동교정 끄기 불가** — API로 완전히 비활성화하는 방법 없음 | \r\n 우회만 가능, 문서 중간에서는 감수해야 함 |
| 6 | **RGBColor BGR 순서** — Windows COLORREF 형식 고정 | 순서를 알고 쓰면 됨, 변경 불가 |

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

// HWPML2X 라운드트립 (양식 채우기)
var xml = hwp.GetTextFile("HWPML2X", "");
// ... XML 파싱 및 수정 ...
hwp.Run("SelectAll");
hwp.Run("Delete");
hwp.SetTextFile(modifiedXml, "HWPML2X", "");

// 위치 읽기/이동
var ps = hwp.GetPosBySet();
var list = ps.Item("List");
var para = ps.Item("Para");
var pos = ps.Item("Pos");

// 페이지별 텍스트 추출
hwp.PageCount;                      // 전체 페이지 수
hwp.GetPageText(0, "UNICODE");      // 1페이지 텍스트 (0-based)
// 포맷 파라미터는 무시됨 — 항상 순수 텍스트 반환

// SelectText로 범위 선택 후 HWPML2X 추출
var offset = getParaOffset(0);
hwp.SelectText(0, offset, 0, offset + 10);
hwp.GetTextFile("HWPML2X", "saveblock");  // 선택 영역만 XML

// 찾아 바꾸기
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
hwp.HParameterSet.HFindReplace.FindString = "old";
hwp.HParameterSet.HFindReplace.ReplaceString = "new";
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
hwp.HParameterSet.HFindReplace.ReplaceMode = 1;
hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
```
