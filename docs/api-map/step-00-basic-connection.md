# 단계 0: 기본 연결

## 1. 공식 API 문서 조사

### 출처

- [한컴 개발자 포털](https://developer.hancom.com/hwpautomation) — HWP Automation Guide (PDF)
- [GitHub: hancom-io/devcenter-archive](https://github.com/hancom-io/devcenter-archive/tree/main/hwp-automation)
- [한컴 개발자 포럼](https://forum.developer.hancom.com)
- 주의: developer.hancom.com/webhwp 쪽은 **HwpCtrl** (웹 ActiveX) 문서이고, **HwpObject** (COM 자동화)와는 별개 API

### HwpCtrl vs HwpObject

|               | HwpCtrl                           | HwpObject                             |
| ------------- | --------------------------------- | ------------------------------------- |
| 정체          | ActiveX 컨트롤 (in-process)       | OLE Automation (out-of-process)       |
| ProgID        | `HWPCtrl.HwpCtrl.1`               | `HWPFrame.HwpObject`                  |
| 용도          | 앱 안에 한글 에디터 임베딩        | 외부에서 독립 프로세스 한글 원격 제어 |
| 호스트 요건   | ActiveX 컨테이너 필수 (VB/MFC/IE) | IDispatch만 있으면 됨                 |
| 우리 프로젝트 | ❌ Node.js/Rust/Electron에서 불가 | ✅ 사용 중                            |
| API 공유      | Action/ParameterSet은 공유        | 프로퍼티/메서드는 각자 다름           |

### 공식 문서 기준 정보 (요약)

- **XHwpWindows.Item(index)**: 0-based, Visible read/write
- **Version**: String, read-only. 10.x=2018, 11.x=2020, 12.x=2022, 13.x=2024
- **EditMode**: I4, read/write. 0=읽기전용, 1=일반, 2=양식, 16=배포(set 불가)
- **IsEmpty**: Bool, read-only
- **PageCount**: I4, read-only. 접근 시 페이지네이션 트리거 (성능 주의)
- **SetMessageBoxMode**: 비트 플래그. 0x10000=자동예, 0xF0000=리셋

---

## 2. ITypeInfo 덤프 vs 공식 문서 차이 → 실험 결과

| 항목                     | 공식 문서                   | ITypeInfo 덤프 | 실험 결과           | 결론                                       |
| ------------------------ | --------------------------- | -------------- | ------------------- | ------------------------------------------ |
| CLSID 프로퍼티           | 존재                        | 없음           | ❌ 0x80020006       | **HwpCtrl 전용이거나 버전 차이**           |
| CurMetatagState          | 존재                        | 없음           | ❌ 0x80020006       | 동일                                       |
| IsTrackChange            | 존재                        | 없음           | ❌ 0x80020006       | 동일                                       |
| IsTrackChangePassword    | 존재                        | 없음           | ❌ 0x80020006       | 동일                                       |
| HParameterSet.HCharShape | ITypeInfo에 429 멤버로 나옴 | get 가능 표시  | ❌ 0x80020006       | **ITypeInfo에는 있지만 런타임 접근 불가!** |
| EditMode put(0)          | 0=읽기전용                  | put 가능       | ⚠️ put(0) → 결과 16 | 0이 아닌 16으로 변경됨 (아래 분석)         |

---

## 3. 테스트 결과 (전체)

### 환경

- 한글 2018 (Version = 10, 0, 0, 5060)
- HWPFrame.HwpObject (COM out-of-process)

### 스칼라 프로퍼티 (get)

| #    | 프로퍼티               | 결과 | 값               | 비고                |
| ---- | ---------------------- | :--: | ---------------- | ------------------- |
| 0-3  | Version                |  ✅  | "10, 0, 0, 5060" | 한글 2018           |
| 0-4  | EditMode               |  ✅  | 1                | 일반 편집 모드      |
| 0-5  | IsEmpty                |  ✅  | true             |                     |
| 0-6  | PageCount              |  ✅  | 1                |                     |
| 0-7  | IsModified             |  ✅  | false            |                     |
| 0-8  | IsPrivateInfoProtected |  ✅  | false            |                     |
| 0-9  | Path                   |  ✅  | "" (빈 문자열)   | 새 문서라 경로 없음 |
| 0-10 | SelectionMode          |  ✅  | 0                |                     |
| 0-11 | CurFieldState          |  ✅  | 0                |                     |

### 공식 문서에만 있는 프로퍼티

| #    | 프로퍼티              | 결과 | 비고                                             |
| ---- | --------------------- | :--: | ------------------------------------------------ |
| 0-12 | CLSID                 |  ❌  | 0x80020006 — HwpCtrl 전용이거나 한글 2018에 없음 |
| 0-13 | CurMetatagState       |  ❌  | 동일                                             |
| 0-14 | IsTrackChange         |  ❌  | 동일                                             |
| 0-15 | IsTrackChangePassword |  ❌  | 동일                                             |

### Dispatch 프로퍼티

| #    | 프로퍼티         |   결과   | 멤버 수 | 비고                        |
| ---- | ---------------- | :------: | :-----: | --------------------------- |
| 0-16 | Application      |    ✅    |   195   | 루트와 동일 (자기 참조)     |
| 0-17 | CharShape        |    ✅    |   21    | ParameterSet 인터페이스     |
| 0-18 | ParaShape        |    ✅    |   21    | ParameterSet 인터페이스     |
| 0-19 | CellShape        | ⚠️ empty |    -    | 표 안이 아니라서            |
| 0-20 | HAction          |    ✅    |   11    | Execute, GetDefault, Run 등 |
| 0-21 | HParameterSet    |    ✅    |   429   | Dispatch 자체는 반환됨      |
| 0-22 | EngineProperties |    ✅    |   21    | ParameterSet                |
| 0-23 | ViewProperties   |    ✅    |   21    | ParameterSet                |
| 0-24 | HeadCtrl         |    ✅    |   16    | 빈 문서에서도 존재          |
| 0-25 | LastCtrl         |    ✅    |   16    | 빈 문서에서도 존재          |
| 0-26 | CurSelectedCtrl  | ⚠️ empty |    -    | 선택 없어서                 |
| 0-27 | ParentCtrl       | ⚠️ empty |    -    | 부모 없어서                 |
| 0-28 | XHwpDocuments    |    ✅    |   16    |                             |
| 0-29 | XHwpWindows      |    ✅    |   14    |                             |
| 0-30 | XHwpMessageBox   |    ✅    |   15    |                             |
| 0-31 | XHwpODBC         |    ✅    |   18    |                             |

### 컬렉션 상세

| #    | 테스트                            | 결과 | 값       |
| ---- | --------------------------------- | :--: | -------- |
| 0-32 | XHwpDocuments.Count               |  ✅  | 1        |
| 0-33 | XHwpWindows.Count                 |  ✅  | 1        |
| 0-34 | XHwpDocuments.Active_XHwpDocument |  ✅  | Dispatch |
| 0-35 | XHwpWindows.Active_XHwpWindow     |  ✅  | Dispatch |

### ParameterSet 상세

| #     | 테스트                           | 결과 | 값          | 비고                     |
| ----- | -------------------------------- | :--: | ----------- | ------------------------ |
| 0-36a | CharShape.SetID                  |  ✅  | "CharShape" |                          |
| 0-36b | CharShape.Item("Height")         |  ✅  | 1000        | 10pt (기본값)            |
| 0-36c | CharShape.Item("Bold")           |  ⚠️  | empty       | 미설정 상태 = empty 반환 |
| 0-37a | ParaShape.SetID                  |  ✅  | "ParaShape" |                          |
| 0-37b | ParaShape.Item("Alignment")      |  ⚠️  | empty       | 미설정 상태              |
| 0-38  | **HParameterSet.HCharShape**     |  ❌  | 0x80020006  | **핵심 발견**            |
| 0-39  | **HParameterSet.HInsertText**    |  ❌  | 0x80020006  | **핵심 발견**            |
| 0-40  | **HParameterSet.HTableCreation** |  ❌  | 0x80020006  | **핵심 발견**            |

### EditMode 쓰기

| #     | 테스트      | 결과 | 비고                                              |
| ----- | ----------- | :--: | ------------------------------------------------- |
| 0-41a | put(0)      |  ⚠️  | EditMode = 0 으로 설정했으나 읽으면 **16**이 나옴 |
| 0-41b | put(1) 복원 |  ✅  | 정상 복원                                         |

### SetMessageBoxMode

| 설정값            | Set 반환 | Get 결과 | 비고               |
| ----------------- | -------- | -------- | ------------------ |
| 0x10000 (자동 예) | 0        | 65536    | 반환값 = 이전 모드 |
| 0x01000 (예 클릭) | 65536    | 4096     |                    |
| 0x02000 (아니오)  | 4096     | 8192     |                    |
| 0x04000 (취소)    | 8192     | 16384    |                    |
| 0xF0000 (리셋)    | 16384    | 983040   |                    |

> Set 반환값 = **이전 모드**, Get = **현재 모드**

### 유틸리티

| #    | 메서드               | 결과 | 값       | 비고                      |
| ---- | -------------------- | :--: | -------- | ------------------------- |
| 0-42 | RGBColor(255,0,0)    |  ✅  | 255      | = 0x0000FF → **BGR 순서** |
| 0-43 | RGBColor(0,0,255)    |  ✅  | 16711680 | = 0xFF0000 → BGR 확인     |
| 0-44 | MiliToHwpUnit(25.4)  |  ✅  | 7200     | 1인치 = 7200 HWP단위      |
| 0-45 | PointToHwpUnit(12.0) |  ✅  | 1200     | 1pt = 100 HWP단위         |
| 0-46 | HwpLineType("Solid") |  ✅  | 1        |                           |
| 0-47 | HAlign("Center")     |  ✅  | 3        |                           |
| 0-48 | TextAlign("Center")  |  ✅  | 2        | HAlign과 값이 다름!       |

### IsActionEnable

| 액션              | 사용 가능 | 비고                                   |
| ----------------- | :-------: | -------------------------------------- |
| InsertText        |  ✅ true  |                                        |
| CharShape         |  ✅ true  |                                        |
| ParagraphShape    |  ✅ true  |                                        |
| TableCreate       |  ✅ true  |                                        |
| AllReplace        |  ✅ true  |                                        |
| PageSetup         |  ✅ true  |                                        |
| InsertPicture     | ❌ false  | CreateAction도 empty 반환이었음 — 연관 |
| Style             | ❌ false  | CreateAction도 empty                   |
| BreakSection      |  ✅ true  |                                        |
| InsertColumnBreak | ❌ false  | CreateAction도 empty                   |
| Delete            |  ✅ true  |                                        |
| Copy              | ❌ false  | 선택 없어서                            |
| Cut               | ❌ false  | 선택 없어서                            |
| Paste             |  ✅ true  |                                        |
| Undo              | ❌ false  | 히스토리 없어서                        |
| Redo              | ❌ false  | 히스토리 없어서                        |
| FileNew           |  ✅ true  |                                        |
| FileOpen          |  ✅ true  |                                        |
| FileSaveAs        |  ✅ true  |                                        |
| Print             |  ✅ true  |                                        |

---

## 4. 분석

### 핵심 발견 1: HParameterSet은 작동한다! (오진 정정)

step-00 초기 테스트에서 "HParameterSet 접근 불가"로 기록했으나, **오진이었음**.
에러는 `hps.get("HCharShape")`가 아니라 그 다음 `d.get("SetID")`에서 발생한 것.

**실제 상황**: HParameterSet.HCharShape는 **ParameterSet이 아니라 별도의 전용 인터페이스**를 가진다.

| 구분               | CreateAction.CreateSet()                             | HParameterSet.HCharShape                        |
| ------------------ | ---------------------------------------------------- | ----------------------------------------------- |
| 인터페이스         | 범용 딕셔너리 (21 멤버)                              | **전용 프로퍼티 (144 멤버)**                    |
| 값 접근            | `set.Item("Height")` / `set.SetItem("Height", 2400)` | `hcs.get("Height")` / `hcs.put("Height", 2400)` |
| SetID/Item/SetItem | ✅ 있음                                              | ❌ 없음 (다른 인터페이스)                       |
| Height, Bold 등    | ❌ 없음                                              | ✅ 직접 프로퍼티로 존재                         |
| HSet 프로퍼티      | ❌ 없음                                              | ✅ HAction에 넘기는 Dispatch                    |

**HParameterSet.HCharShape 주요 프로퍼티** (144개 중 발췌):

- `Height` (I4, get/put) — 글자 크기
- `Bold`, `Italic` (VT(18), get/put) — 볼드/이탤릭
- `TextColor`, `UnderlineColor`, `StrikeOutColor` (VT(19), get/put) — 색상
- `FaceNameHangul`, `FaceNameLatin` 등 (String, get/put) — 글꼴
- `HSet` (Dispatch, get/put) — HAction에 넘기는 ParameterSet

**HParameterSet.HInsertText** (12 멤버):

- `Text` (String, get/put) — 삽입할 텍스트
- `HSet` (Dispatch, get/put)

**결론**: 두 패턴 모두 사용 가능. HParameterSet은 직접 프로퍼티 방식이라 타입이 명확하고 프로퍼티 목록을 알 수 있다는 장점이 있음.

### 핵심 발견 2: EditMode put(0) → 16

공식 문서: EditMode=0은 "읽기전용", EditMode=16은 "배포문서(읽기전용, set 불가)".
하지만 put(0)을 했더니 EditMode가 **16**으로 읽힘.

가능한 해석:

- 한글 2018에서는 0을 설정하면 내부적으로 16(배포모드)으로 변환되는 것일 수 있음
- 또는 get이 0 대신 16을 반환하는 매핑일 수 있음
- 어쨌든 put(1)로 정상 복원은 됨

### 핵심 발견 3: RGBColor는 BGR 순서

`RGBColor(R, G, B)` 반환값은 Windows COLORREF 형식 (`0x00BBGGRR`):

- RGBColor(255, 0, 0) = 255 = 0x000000FF (빨강 → RR 위치)
- RGBColor(0, 0, 255) = 16711680 = 0x00FF0000 (파랑 → BB 위치)

### 핵심 발견 4: HAlign vs TextAlign

둘 다 "Center"를 넣었는데 값이 다름 (HAlign=3, TextAlign=2). 이름은 비슷하지만 **다른 열거형 체계**. 사용 시 혼동 주의.

### 핵심 발견 5: InsertPicture는 IsActionEnable=false

빈 문서 상태에서 InsertPicture가 false. CreateAction("InsertPicture")도 empty를 반환했음.

- 문서 조건이 필요한지, 아니면 보안모듈 관련인지 → 이후 단계에서 확인 필요

### 핵심 발견 6: CharShape.Item("Bold") = empty

Item("Height")는 1000을 반환하지만, Item("Bold")는 empty.

- "설정되지 않은" 상태를 empty로 표현하는 것으로 보임
- Height는 항상 기본값이 있지만, Bold는 명시적 설정 전까지 empty

### 기타 발견

- `SetMessageBoxMode` 반환값 = 이전 모드, `GetMessageBoxMode` = 현재 모드
- HeadCtrl/LastCtrl은 빈 문서에서도 Dispatch 반환 (문서 구조의 루트 컨트롤?)
- Copy/Cut은 선택이 없으면 false, Undo/Redo는 히스토리 없으면 false (상태 의존적)
- 리셋(0xF0000) 후 Get이 983040 (= 0xF0000) → 리셋은 "0으로 돌리기"가 아니라 "리셋 플래그를 모드로 설정"

---

## 5. 최종 결론

### 정상 작동 (✅)

- COM 연결, 창 표시 — 문제없음
- 스칼라 프로퍼티 9개 전부 정상
- Dispatch 프로퍼티 13/16개 정상 (나머지 3개는 빈 문서 상태라 empty — 정상)
- 컬렉션 접근 (Count, Active 등) — 정상
- ParameterSet 읽기 (SetID, Item) — 정상 (empty는 미설정 의미)
- EditMode 쓰기 — 동작함 (값 매핑 주의)
- SetMessageBoxMode 전 모드 — 정상
- 유틸리티 메서드 (RGBColor, 단위변환, 열거형) — 전부 정상
- IsActionEnable — 상태에 따라 정확히 반영

### 사용 불가 (❌) — 중요하지 않음

| 프로퍼티 | 용도 | 영향 |
|----------|------|------|
| CLSID | COM 클래스 식별자 조회 | 우리가 쓸 일 없음 |
| CurMetatagState | 메타태그 상태 | 문서 자동화와 무관 |
| IsTrackChange | 변경 추적 모드 여부 | 문서 편집에 필요 없음 |
| IsTrackChangePassword | 변경 추적 비밀번호 | 동일 |

### 오진 정정 (⚠️→✅)

- **HParameterSet 하위 접근** — 초기에 불가로 기록했으나, 실제로는 **접근 가능**. `.SetID`/`.Item`이 아니라 `.Height`/`.Bold` 같은 **직접 프로퍼티**를 써야 함. 상세는 "핵심 발견 1" 참조.
- **JS Proxy 경유 테스트 완료** — 공식 문서 문법 (`hwp.HParameterSet.HInsertText.Text = "..."`) 이 그대로 작동함을 확인.

### 주의 필요 (⚠️) — 이후 단계에서 추가 검증 예정

| 항목 | 현상 | 검증 단계 |
|------|------|-----------|
| EditMode put(0) → 16 | 0으로 설정했는데 읽으면 16 | 단계 1 (문서 생명주기) |
| RGBColor BGR 순서 | R,G,B 인자 → 반환값은 BGR | 단계 5 (글자 서식) |
| HAlign ≠ TextAlign | 같은 "Center"인데 값이 다름 (3 vs 2) | 단계 6 (문단 서식) |
| InsertPicture 비활성 | 빈 문서에서 IsActionEnable=false | 단계 8 (그림 삽입) |
| CharShape.Item("Bold") = empty | 미설정 항목은 empty 반환 | 단계 5 (글자 서식) |
| CellShape = empty | 표 안이 아니면 empty | 단계 7 (표) |
| CurSelectedCtrl = empty | 선택 없으면 empty | 단계 3 (커서/선택) |

### 다음 단계

단계 1 (문서 생명주기: Open/Save/Close)로 진행.
