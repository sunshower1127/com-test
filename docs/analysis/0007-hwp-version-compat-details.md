# HWP 버전별 COM 연결 상세 분석

레퍼런스 `0003-hwp호환_대화내역.md`와 웹 검색 결과를 종합한 실무 분석.

---

## 1. ProgID / CLSID 체계

### ProgID 3단계 구조

| ProgID | 의미 | 연결 대상 |
|--------|------|----------|
| `HWPFrame.HwpObject` | 버전 독립 | CurVer이 가리키는 최신 버전 |
| `HWPFrame.HwpObject.1` | 한글 2020 이하 | CLSID `{2291CEFF-...}` |
| `HWPFrame.HwpObject.2` | 한글 2022 이상 | CLSID `{2291CF00-...}` |

```
HKCR\HWPFrame.HwpObject\CurVer → "HWPFrame.HwpObject.2"  (마지막 설치 버전)
```

### CLSID 차이

| 버전 | CLSID | 차이점 |
|------|-------|--------|
| 2018, 2020 | `{2291CEFF-64A1-4877-A9B4-68CFE89612D6}` | 마지막 바이트 `CEFF` |
| 2022, 2024 | `{2291CF00-64A1-4877-A9B4-68CFE89612D6}` | 마지막 바이트 `CF00` |

TypeLib GUID는 전 버전 공통: `{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}`

### 우리 코드 영향

- `DispatchObject::create("HWPFrame.HwpObject")` → ProgID 기반 → **CLSID 차이 자동 해결** ✅
- `find_install_path()`에서 `{2291CF00-...}` 하드코딩 → **2018/2020에서 경로 못 찾음** ❌
- **수정 필요**: 양쪽 CLSID 모두 탐색하거나 ProgID → CLSID 동적 조회

---

## 2. 내부 버전 번호 체계

| 한글 버전 | 내부 버전 | 코드명 | 설정 경로 |
|-----------|----------|--------|----------|
| 2018 | 10.x | 100 | `%appdata%\HNC\User\Common\100` |
| 2020 | 11.x | 110 | `%appdata%\HNC\User\Common\110` |
| 2022 | 12.x | 120 | `%appdata%\HNC\User\Common\120` |
| 2024 | 13.x | 130 | `%appdata%\HNC\User\Common\130` |

런타임에 `hwp.Version` 프로퍼티로 확인 가능.

---

## 3. 다중 버전 설치 충돌

`HWPFrame.HwpObject` ProgID는 **하나의 LocalServer32 경로만 가리킬 수 있음**.
→ 마지막으로 설치/업데이트된 버전이 COM 서버를 독점.

### 특정 버전으로 전환

```batch
:: 관리자 권한 필요
"C:\Program Files (x86)\Hnc\Office 2018\HOffice100\Bin\Hwp.exe" -regserver  :: 2018로 전환
"C:\Program Files (x86)\Hnc\Office 2024\HOffice130\Bin\Hwp.exe" -regserver  :: 2024로 전환
```

### `HWPFrame.HwpObject.1` vs `.2`로 분기 가능?

이론적으로 별도 CLSID에 연결하여 동시 사용 가능하지만,
실제로는 두 ProgID가 같은 CLSID를 가리키는 경우가 대부분.
**레지스트리 수동 편집 없이는 신뢰하기 어려움**.

### 우리 코드 대응

런타임에 `hwp.Version`으로 연결된 버전 확인하고 로그에 출력하는 것이 안전.

---

## 4. 한글 2024 Breaking Changes 상세

### 4-1. Visible 속성 버그

**증상**: `Visible = True` 해도 작업표시줄 미리보기만 나타남

**워크어라운드** (한컴 공식 확인):
```python
hwp.XHwpWindows.Item(0).Visible = False  # 먼저 끔
hwp.XHwpWindows.Item(0).Visible = True   # 다시 켬
```

**우리 코드 대응**:
```rust
// hwp-bridge set_visible에서 버전 감지 후 토글
pub fn set_visible(&self, visible: bool) -> Result<()> {
    let windows = self.disp.get("XHwpWindows")?.into_dispatch()?;
    let window = windows.get_by("Item", &[Variant::I32(0)])?.into_dispatch()?;
    if visible {
        // 2024 워크어라운드: False→True 토글
        let _ = window.put("Visible", false);
    }
    window.put("Visible", visible)
}
```

### 4-2. .NET Interop RPC 에러

**증상**: C# Interop DLL 통한 COM 호출 시 "Server RPC 에러"
- 특히 `hwp.HParameterSet.HParaShape.BorderFill` 같은 중첩 프로퍼티

**원인**: early-binding (Interop DLL 타입 참조)이 2024에서 깨짐

**해결**: `dynamic` 키워드로 late-binding 사용

**우리 코드**: Rust IDispatch late-binding 사용 → **영향 없음** ✅

### 4-3. 다중 인스턴스 충돌

**증상**: 두 번째 인스턴스 생성 시 첫 번째 창 사라짐

**상태**: 2025년 2월 패치로 수정됨

**우리 코드 대응**: 동시에 2개 이상 HWP 인스턴스 생성 지양. 1개만 사용.

### 4-4. RegisterModule 팝업 지속

**증상**: `RegisterModule` 반환값 True인데도 보안 팝업 계속 표시

**체크리스트**:
1. 레지스트리 경로가 `HwpAutomation\Modules`인지 (HwpCtrl 아닌지)
2. 값 이름이 `RegisterModule()` 두 번째 인자와 정확히 일치하는지
3. DLL 경로에 따옴표(`"`) 없는지
4. 최신 패치 적용했는지
5. `hwp.exe -regserver` 실행했는지

### 4-5. 패치 후 COM 등록 깨짐

**증상**: 패치 적용 후 OLE 기능 예외, 매크로 실행 실패

**해결**:
```batch
:: 관리자 권한
"C:\Program Files (x86)\Hnc\Office 2024\HOffice130\Bin\Hwp.exe" /RegServer
```

**빈도**: 한글 2024는 패치마다 발생할 수 있음. 2018~2022에서는 드묾.

### 4-6. 2024.06 패치 — 구버전 OCX 소급 삭제

**충격적 변경**: 한글 2024뿐 아니라 2018/2020/2022에 설치된 OCX 파일도 삭제.
- 우리는 HwpAutomation만 사용 → **영향 없음** ✅
- HwpCtrl 기반 레거시 시스템은 이 패치 이후 동작 불가

---

## 5. 성능 최적화 팁

### Visible=False 백그라운드 실행

대량 작업 시 **체감 10배 이상 속도 향상**:
```python
hwp.XHwpWindows.Item(0).Visible = False  # 화면 갱신 끄기
# ... 대량 작업 ...
hwp.XHwpWindows.Item(0).Visible = True   # 완료 후 켜기
```

### PutFieldText 배치 입력

커서 이동 + InsertText 반복 대신, 누름틀(필드)에 일괄 입력:
```python
hwp.PutFieldText("필드1\x02필드2\x02필드3", "값1\x02값2\x02값3")
```

### CreateAction 캐싱

`CreateAction` 결과를 변수에 저장하여 재사용:
```js
// ✅ 한 번 생성, 여러 번 사용
var act = hwp.CreateAction("InsertText")
var set = act.CreateSet()

for (var i = 0; i < 100; i++) {
  act.GetDefault(set)
  set.SetItem("Text", "line " + i)
  act.Execute(set)
}
```

### ReleaseScan 필수

`InitScan/GetText` 사용 후 **반드시 `ReleaseScan()` 호출**. 미호출 시 메모리 누수.

---

## 6. TypeLib 참고사항

한글 설치 디렉터리 `Bin` 폴더에 두 개의 `.tlb` 파일 존재:
- `HwpObject.tlb` — **권장** (한컴 공식)
- `HwpAutomation.tlb` — 부모 윈도우 없이는 제대로 동작 안 함

> "HwpAutomation"은 ProgID가 아니라 타입 라이브러리 이름.
> COM 연결에는 반드시 `HWPFrame.HwpObject` 사용.

---

## 7. 우리에게 유리한 점

| 항목 | 이유 |
|------|------|
| late-binding | Rust IDispatch 기반 → .NET Interop RPC 에러 무관 |
| out-of-process COM | 32/64bit 자동 마샬링 |
| ProgID 기반 연결 | CLSID 차이 자동 해결 |
| Drop trait에서 Quit() | 좀비 프로세스 방지 |
| CreateAction 패턴 | HParameterSet 직접 put보다 안정적 |

---

## 출처

- `docs/references/0003-hwp호환_대화내역.md`
- https://forum.developer.hancom.com/t/2024-hwpctrl/1013
- https://forum.developer.hancom.com/t/topic/1715
- https://forum.developer.hancom.com/t/2024-2/2235
- https://forum.developer.hancom.com/t/2024-registermodule/1655
- https://forum.developer.hancom.com/t/topic/2192
- https://forum.developer.hancom.com/t/250624-2024/2616
- https://pyhwpx.com/67
- https://pyhwpx.com/30
- https://eduhub.co.kr/board_vRlj10/182153
