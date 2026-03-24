# HWP COM API 목록 (Raw)

> 출처: `hwp_api_dump` 예제 실행 결과 (ITypeInfo 기반)
> 날짜: 2026-03-24

## 루트 객체: HWPFrame.HwpObject

총 195 멤버 (IUnknown/IDispatch 인프라 제외 시 실질 ~165개)

### 메서드 (실질적으로 사용할 수 있는 것만)

COM 인프라 메서드(AddRef, Release, QueryInterface, GetIDsOfNames, GetTypeInfo, GetTypeInfoCount, Invoke)는 제외.

#### 핵심 문서 조작

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| Open | (filename: String, Format: Variant, arg: Variant) | Bool | 문서 열기 |
| Save | (save_if_dirty: Variant) | Bool | 저장 |
| SaveAs | (Path: String, Format: Variant, arg: Variant) | Bool | 다른 이름으로 저장 |
| Clear | (option: Variant) | void | 문서 닫기 |
| Quit | () | void | 한글 종료 |
| Insert | (Path: String, Format: Variant, arg: Variant) | void | 파일 삽입 |

#### 텍스트 입출력

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| GetTextFile | (Format: String, option: String) | Variant | 전체 텍스트 추출 |
| SetTextFile | (data: Variant, Format: String, option: String) | I4 | 텍스트 설정 |
| GetText | (Text: Ptr) | I4 | 텍스트 가져오기 (Ptr 사용) |
| GetPageText | (pgno: I4, option: Variant) | String | 특정 페이지 텍스트 |
| GetHeadingString | () | String | 제목 문자열 |

#### Action/Set 시스템

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| CreateAction | (actidstr: String) | Dispatch | Action 객체 생성 |
| CreateSet | (setidstr: String) | Dispatch | ParameterSet 객체 생성 |
| Run | (ActID: String) | void | 파라미터 없는 액션 실행 |
| IsActionEnable | (actionID: String) | Bool | 액션 실행 가능 여부 |
| ReplaceAction | (OldActionID: String, NewActionID: String) | Bool | 액션 교체 |
| InitHParameterSet | () | void | HParameterSet 초기화 |

#### 커서/위치

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| MovePos | (moveID: Variant, Para: Variant, pos: Variant) | Bool | 커서 이동 |
| SetPos | (List: I4, Para: I4, pos: I4) | Bool | 위치 설정 |
| GetPos | (List: Ptr, Para: Ptr, pos: Ptr) | void | 위치 가져오기 (Ptr) |
| GetPosBySet | () | Dispatch | 위치를 Set으로 |
| SetPosBySet | (dispVal: Dispatch) | Bool | Set으로 위치 설정 |
| SelectText | (spara: I4, spos: I4, epara: I4, epos: I4) | Bool | 텍스트 영역 선택 |
| GetSelectedPos | (6 Ptr params) | Bool | 선택 영역 위치 (Ptr) |
| GetSelectedPosBySet | (sset: Dispatch, eset: Dispatch) | Bool | 선택 영역 위치 (Set) |

#### 필드

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| CreateField | (Direction: String, memo: Variant, name: Variant) | Bool | 필드 생성 |
| FieldExist | (Field: String) | Bool | 필드 존재 확인 |
| GetFieldList | (Number: Variant, option: Variant) | String | 필드 목록 |
| GetFieldText | (Field: String) | String | 필드 텍스트 |
| PutFieldText | (Field: String, Text: String) | void | 필드 텍스트 설정 |
| GetCurFieldName | (option: Variant) | String | 현재 필드 이름 |
| SetCurFieldName | (Field: String, option: Variant, Direction: String, memo: String) | Bool | 필드 이름 설정 |
| MoveToField | (Field: String, Text: Variant, start: Variant, select: Variant) | Bool | 필드로 이동 |
| ModifyFieldProperties | (Field: String, remove: I4, Add: I4) | I4 | 필드 속성 수정 |
| RenameField | (oldname: String, newname: String) | void | 필드 이름 변경 |
| SetFieldViewOption | (option: I4) | I4 | 필드 표시 옵션 |

#### 컨트롤(그림/OLE/객체)

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| InsertPicture | (Path + 7 Variant params) | Dispatch | 그림 삽입 |
| InsertBackgroundPicture | (BorderType + 7 params) | Bool | 배경 그림 삽입 |
| InsertCtrl | (CtrlID: String, initparam: Variant) | Dispatch | 컨트롤 삽입 |
| DeleteCtrl | (ctrl: Dispatch) | Bool | 컨트롤 삭제 |
| FindCtrl | () | String | 컨트롤 찾기 |
| UnSelectCtrl | () | void | 컨트롤 선택 해제 |
| CheckXObject | (bstring: String) | Dispatch | XObject 확인 |

#### 스캔(반복 탐색)

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| InitScan | (option + 5 Variant params) | Bool | 스캔 시작 |
| GetText | (Text: Ptr) | I4 | 다음 텍스트 가져오기 |
| ReleaseScan | () | void | 스캔 종료 |

#### 이미지/인쇄

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| CreatePageImage | (Path + 4 Variant params) | Bool | 페이지 이미지 생성 |
| SetBarCodeImage | (7 params) | Bool | 바코드 이미지 설정 |
| GetBinDataPath | (binid: I4) | String | 바이너리 데이터 경로 |

#### 유틸리티/변환

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| RGBColor | (red, green, blue: VT(17)) | I4 | RGB 색상값 생성 |
| MiliToHwpUnit | (mili: F8) | I4 | mm → HWP단위 |
| PointToHwpUnit | (Point: F8) | I4 | pt → HWP단위 |
| LunarToSolar / SolarToLunar | (날짜 params) | Bool | 음양력 변환 |
| LunarToSolarBySet / SolarToLunarBySet | (날짜 params) | Dispatch | 음양력 변환 (Set) |
| ConvertPUAHangulToUnicode | (Text: Variant) | I4 | PUA한글→유니코드 |

#### 보안/DRM

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| RegisterModule | (ModuleType: String, ModuleData: Variant) | Bool | 보안 모듈 등록 |
| SetDRMAuthority | (authority: I4) | Bool | DRM 권한 설정 |
| FindPrivateInfo | (PrivateType: I4, PrivateString: Variant) | I4 | 개인정보 찾기 |
| ProtectPrivateInfo | (PotectingChar: String, ...) | Bool | 개인정보 보호 |
| RegisterPrivateInfoPattern | (PrivateType: I4, PrivatePattern: String) | Bool | 패턴 등록 |
| SetPrivateInfoPassword | (Password: String) | Bool | 비밀번호 설정 |

#### 기타

| 메서드 | 시그니처 | 반환 | 설명 (추정) |
|--------|----------|------|-------------|
| SetMessageBoxMode | (Mode: I4) | I4 | 메시지박스 모드 (팝업 억제) |
| GetMessageBoxMode | () | I4 | 현재 메시지박스 모드 |
| LockCommand | (ActID: String, isLock: Bool) | void | 명령 잠금 |
| IsCommandLock | (actionID: String) | Bool | 명령 잠금 여부 |
| KeyIndicator | (8 Ptr params) | Bool | 상태표시줄 정보 |
| GetMousePos | (XRelTo: I4, YRelTo: I4) | Dispatch | 마우스 위치 |
| SetTitleName | (Title: String) | void | 타이틀 설정 |
| GetFileInfo | (filename: String) | Dispatch | 파일 정보 |
| GetScriptSource | (filename: String) | String | 스크립트 소스 |
| RunScriptMacro | (FunctionName + 2 params) | Bool | 스크립트 매크로 실행 |
| ExportStyle / ImportStyle | (param: Dispatch) | Bool | 스타일 내보내기/가져오기 |

#### 열거형 변환 메서드 (String → I4)

`SomeEnum("StringValue") → I4` 패턴. 열거형 문자열을 정수로 변환하는 유틸리티.

ArcType, AutoNumType, BorderShape, BreakWordLatin, BrushType, Canonical, CellApply, CharShadowType, ColDefType, ColLayoutType, CreateID, CreateMode, CrookedSlash, DSMark, DbfCodeType, Delimiter, DrawAspect, DrawFillImage, DrawShadowType, Encrypt, EndSize, EndStyle, FillAreaType, FindDir, FontType, Gradation, GridMethod, GridViewLine, GutterMethod, HAlign, Handler, Hash, HatchStyle, HeadType, HeightRel, Hiding, HorzRel, HwpLineType, HwpLineWidth, HwpOutlineStyle, HwpOutlineType, HwpUnderlineShape, HwpUnderlineType, HwpZoomType, ImageFormat, LineSpacingMethod, LineWrapType, MacroState, MailType, NumberFormat, Numbering, PageNumPosition, PageType, ParaHeadAlign, PicEffect, PlacementType, PresentEffect, PrintDevice, PrintPaper, PrintRange, PrintType, Revision, SideType, Signature, Slash, SortDelimiter, StrikeOut, StyleType, SubtPos, TableBreak, TableFormat, TableSwapType, TableTarget, TextAlign, TextArtAlign, TextDir, TextFlowType, TextWrapType, VAlign, VertRel, ViewFlag, WatermarkBrush, WidthRel

---

### 프로퍼티 (Get)

| 프로퍼티 | 반환 | 설명 (추정) |
|----------|------|-------------|
| Application | Dispatch | 자기 자신 (195 멤버 동일) |
| CellShape | Dispatch | 현재 셀 모양 (빈 문서에서 empty) |
| CharShape | Dispatch | 현재 글자 모양 (ParameterSet) |
| CurFieldState | I4 | 현재 필드 상태 |
| CurSelectedCtrl | Dispatch | 현재 선택된 컨트롤 (없으면 empty) |
| EditMode | I4 | 편집 모드 |
| EngineProperties | Dispatch | 엔진 속성 (ParameterSet) |
| HAction | Dispatch | HAction 객체 (Execute, GetDefault, Run 등) |
| HParameterSet | Dispatch | HParameterSet (429 멤버 — 모든 파라미터셋 집합) |
| HeadCtrl | Dispatch | 문서 첫 컨트롤 (Ctrl 순회용) |
| IsEmpty | Bool | 빈 문서 여부 |
| IsModified | Bool | 수정 여부 |
| IsPrivateInfoProtected | Bool | 개인정보 보호 여부 |
| LastCtrl | Dispatch | 문서 마지막 컨트롤 |
| PageCount | I4 | 페이지 수 |
| ParaShape | Dispatch | 현재 문단 모양 (ParameterSet) |
| ParentCtrl | Dispatch | 부모 컨트롤 (없으면 empty) |
| Path | String | 현재 문서 경로 |
| SelectionMode | I4 | 선택 모드 |
| Version | String | 버전 정보 |
| ViewProperties | Dispatch | 보기 속성 (ParameterSet) |
| XHwpDocuments | Dispatch | 문서 컬렉션 |
| XHwpMessageBox | Dispatch | 메시지박스 객체 |
| XHwpODBC | Dispatch | ODBC 객체 |
| XHwpWindows | Dispatch | 창 컬렉션 |

### 프로퍼티 (Put)

| 프로퍼티 | 타입 | 설명 (추정) |
|----------|------|-------------|
| CellShape | Dispatch | 셀 모양 설정 |
| CharShape | Dispatch | 글자 모양 설정 |
| EditMode | I4 | 편집 모드 설정 |
| EngineProperties | Dispatch | 엔진 속성 설정 |
| ParaShape | Dispatch | 문단 모양 설정 |
| ViewProperties | Dispatch | 보기 속성 설정 |

---

## 하위 객체

### hwp.Application
루트 객체와 동일 (195 멤버). 자기 참조.

### hwp.CharShape (ParameterSet 인터페이스)
21 멤버. `Item()`, `SetItem()`, `ItemExist()` 등으로 속성 조회/설정.

### hwp.ParaShape (ParameterSet 인터페이스)
21 멤버. CharShape와 동일한 인터페이스.

### hwp.EngineProperties (ParameterSet 인터페이스)
21 멤버. 동일 인터페이스.

### hwp.ViewProperties (ParameterSet 인터페이스)
21 멤버. 동일 인터페이스.

### hwp.HAction
11 멤버 (인프라 제외 시 4개):
- `Execute(actname: String, pVal: Dispatch) → Bool`
- `GetDefault(actname: String, pVal: Dispatch) → Bool`
- `PopupDialog(actname: String, pVal: Dispatch) → Bool`
- `Run(actname: String) → Bool`

### hwp.HParameterSet
**429 멤버** — 모든 H-파라미터셋의 집합. 각각이 Dispatch를 반환 (get/put).
예: HCharShape, HParaShape, HInsertText, HTableCreation, HFindReplace, HPageDef 등 약 140+ 종류.

### hwp.HeadCtrl / hwp.LastCtrl (Ctrl 인터페이스)
16 멤버 (인프라 제외 시):
- `GetAnchorPos(type: I4) → Dispatch`
- `CtrlCh → I4`, `CtrlID → String`, `HasList → Bool`
- `Next → Dispatch`, `Prev → Dispatch` (연결 리스트 순회)
- `Properties → Dispatch` (get/put)
- `UserDesc → String`

### hwp.XHwpDocuments
16 멤버:
- `Add(isTab: Bool) → Dispatch` — 새 문서
- `Close(isDirty: Bool)` — 문서 닫기
- `FindItem(lDocID: I4) → Dispatch`
- `Active_XHwpDocument → Dispatch`
- `Count → I4`, `Item(index: I4) → Dispatch`

### hwp.XHwpWindows
14 멤버:
- `Add → Dispatch` — 새 창
- `Close(isDirty: Bool) → Bool`
- `Active_XHwpWindow → Dispatch`
- `Count → I4`, `Item(index: I4) → Dispatch`

### hwp.XHwpMessageBox
15 멤버:
- `DoModal` — 메시지박스 표시
- `string → String` (get/put), `Flag → VT(18)` (get/put), `Result → VT(18)` (get/put)

### hwp.XHwpODBC
18 멤버:
- `Connect`, `Disconnect`, `QueryExcute`, `ResultValue`
- `GetDataByName`, `GetDataByNum`, `Select`, `Insert`, `UpDate`, `Delete`

---

## CreateAction 탐색 결과

| Action ID | CreateAction | CreateSet |
|-----------|-------------|-----------|
| InsertText | ✅ Dispatch | ✅ ParameterSet |
| CharShape | ✅ Dispatch | ✅ ParameterSet |
| ParagraphShape | ✅ Dispatch | ✅ ParameterSet |
| TableCreate | ✅ Dispatch | ✅ ParameterSet |
| AllReplace | ✅ Dispatch | ✅ ParameterSet |
| PageSetup | ✅ Dispatch | ✅ ParameterSet |
| InsertPicture | ❌ empty | - |
| Style | ❌ empty | - |
| BreakSection | ✅ Dispatch | ❌ empty (CreateSet 반환 empty) |
| InsertColumnBreak | ❌ empty | - |

### Action 객체 공통 인터페이스 (인프라 제외)
- `CreateSet → Dispatch` — ParameterSet 생성
- `Execute(param: Dispatch) → Bool` — 실행
- `GetDefault(param: Dispatch) → I4` — 기본값 로드
- `PopupDialog(param: Dispatch) → I4` — 대화상자 표시
- `Run → I4` — 파라미터 없이 실행
- `ActID → String`, `SetID → String`

### ParameterSet 공통 인터페이스 (인프라 제외)
- `Item(itemid: String) → Variant` — 항목 조회
- `SetItem(itemid: String, newVal: Variant)` — 항목 설정
- `ItemExist(itemid: String) → Bool` — 항목 존재 확인
- `Clone → Dispatch`, `Merge(srcset) → Bool`, `GetIntersection(srcset)`
- `CreateItemArray(itemid, Count) → Dispatch`, `CreateItemSet(itemid, SetID) → Dispatch`
- `RemoveItem(itemid)`, `RemoveAll(SetID)`
- `Count → I4`, `IsSet → Bool`, `SetID → String`

---

## 다음 단계

1-2. **카테고리별 실제 호출 시도** — 위 API를 하나씩 실행하고 결과를 사용자가 확인
