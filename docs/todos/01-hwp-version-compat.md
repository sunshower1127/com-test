# HWP 버전별 대응 분석

## 상태: ✅ 조사 완료

## 결론 요약

우리는 **HwpAutomation (`HWPFrame.HwpObject`)만 사용** → **한글 2018 이상만 지원**.
한글 2024는 **대규모 breaking change**가 있어 주의 필요.

| | HwpCtrl (ActiveX) | HwpAutomation (우리) |
|---|---|---|
| ProgID | `HwpCtrl.HwpObject` | `HWPFrame.HwpObject` |
| 방식 | ActiveX (in-process) | COM LocalServer (out-of-process) |
| 32/64bit | 반드시 일치 필요 | 자동 마샬링 (불일치 OK) |
| 지원 버전 | 97 ~ 2022 | **2018 ~ 현재** |
| 2024에서 | 제거됨 | 유일한 방법 |
| 보안 모듈 경로 | `HNC\HwpCtrl\Modules` | `HNC\HwpAutomation\Modules` |

> **참고**: 전통적으로 32비트 전용. 2024부터 64비트 설치 옵션 존재 가능 (확인 필요).

## CLSID / ProgID 버전별 차이 (중요!)

| 항목 | 2018 / 2020 | 2022 / 2024 |
|------|-------------|-------------|
| 버전별 ProgID | `HWPFrame.HwpObject.1` | `HWPFrame.HwpObject.2` |
| CLSID | `{2291CEFF-64A1-4877-...}` | `{2291CF00-64A1-4877-...}` |
| 내부 버전 | 10.x (2018), 11.x (2020) | 12.x (2022), 13.x/130 (2024) |

- `HWPFrame.HwpObject` (버전 독립 ProgID)의 `CurVer` → 마지막 설치된 버전을 가리킴
- **다중 버전 설치 시**: 마지막으로 `/RegServer` 실행한 버전만 COM에 연결됨
- TypeLib GUID는 전 버전 공통: `{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}`

> ⚠️ 우리 코드에서 `{2291CF00-...}`를 하드코딩하고 있음 → 2018/2020에서는 `{2291CEFF-...}`이므로
> **CLSID 하드코딩 대신 ProgID 기반 연결을 사용해야 안전** (현재 `DispatchObject::create("HWPFrame.HwpObject")` → OK)

> ⚠️ `HwpAutomation`은 ProgID가 **아니라** 타입 라이브러리 이름임. 실제 COM 연결은 반드시 `HWPFrame.HwpObject` 사용.

## 버전별 호환성 매트릭스

| 항목 | 2018 | 2020 | 2022 | 2024 |
|------|------|------|------|------|
| HwpAutomation | ✅ | ✅ | ✅ | ✅ |
| HwpCtrl (ActiveX) | ✅ | ✅ | ✅ (deprecated) | ❌ 제거 |
| 다중 인스턴스 | ✅ | ✅ | ✅ | ❌→⚠️ (패치 후 수정됨) |
| Visible 제어 | ✅ | ✅ | ✅ | ⚠️ 불안정 |
| RegisterModule | ✅ | ✅ | ✅ | ⚠️ 팝업 지속 사례 |
| /RegServer 필요 | 드묾 | 드묾 | 드묾 | **패치마다 필요** |
| OCX 소급 삭제 | - | - | - | 2024.06 패치가 구버전 OCX도 삭제 |

---

## 한글 2018

**가장 안정적**. 우리 개발/테스트 환경.

- HwpAutomation 최초 도입
- HwpCtrl (ActiveX) + HwpAutomation 모두 지원
- RegisterModule 정상 동작

**알려진 이슈:**
- C#에서 `NullReferenceException` 사례 (같은 코드가 2020/2022에서는 동작)
- 좀비 프로세스: `Quit()` 호출 없이 dispatch 해제하면 백그라운드에 남음

---

## 한글 2020

**2018과 거의 동일**. 특이사항 없음.

- AI 기능(OfficeTalk) 추가, COM API 변경 없음
- RegisterModule 정상 동작

**알려진 이슈:**
- 업데이트 후 "Server RPC 에러" 사례 (2024에서도 동일 보고)
- HwpCtrl.ocx 등록 실패 사례 → `regsvr32`로 재등록

---

## 한글 2022

**과도기**. OCX 지원 종료 예고.

- 다중 인스턴스 정상 동작
- HwpCtrl + HwpAutomation 모두 지원 (마지막 버전)
- Visible 제어 정상

**알려진 이슈:**
- C++ 32비트에서 메모리 누수 보고 — 자동화 반복 시 HWP 프로세스 누적
  - 해결: `Quit()` → dispatch 해제 → `OleUninitialize()` 순서 엄수

---

## 한글 2024 ⚠️ (대규모 Breaking Changes)

### Breaking Change 1: ActiveX/OCX 완전 제거

> "한컴오피스 2024부터는 한글 컨트롤을 지원하지 않고 있습니다" — 한컴 공식

- HwpCtrl ActiveX OCX 완전 삭제
- **2024.06.24 패치**: 구버전(2018/2020/2022)의 OCX 파일도 소급 삭제!
- 우리는 HwpAutomation만 사용하므로 **직접 영향 없음**

### Breaking Change 2: 다중 인스턴스 깨짐 (패치로 수정)

- 여러 `HWPFrame.HwpObject` 인스턴스 생성 시 마지막 창만 보임
- 새 인스턴스가 별도 창이 아닌 같은 창 안에 생성됨
- **수정**: 이후 패치에서 수정됨. 최신 패치 적용 필요

### Breaking Change 3: Visible 속성 불안정

- `hwp.XHwpWindows.Item(0).Visible = True` → 작업표시줄 미리보기만 나오고 실제 창 안 보임
- **워크어라운드**: 토글 2번
  ```
  hwp.XHwpWindows.Item(0).Visible = False
  hwp.XHwpWindows.Item(0).Visible = True
  ```
- "최근 문서 열기" 설정과 관련

### Breaking Change 3.5: .NET Interop RPC 에러

- C# Interop DLL 통한 COM 호출 시 "Server RPC 에러" 발생
- `hwp.HParameterSet.HParaShape.BorderFill` 같은 중첩 프로퍼티 접근에서 특히 문제
- **해결**: `dynamic` 키워드로 late-binding 사용 (early-binding 대신)
- **우리 코드**: Rust에서 IDispatch late-binding 사용 중이므로 **영향 없음** ✅

### Breaking Change 4: RegisterModule 팝업 지속

- `RegisterModule` 반환값 `True`인데도 보안 팝업이 계속 표시됨
- 반드시 **Object 버전** 모듈 사용 (ActiveX 버전 ≠ Object 버전)
- 레지스트리 값 이름이 `RegisterModule()` 두 번째 인자와 정확히 일치해야 함
- 포럼에서 완전한 해결 미확인

### Breaking Change 5: 패치 후 COM 등록 깨짐

- 패치 적용 후 OLE 기능 예외, 커스텀 메뉴/도구 상자 미표시
- 여러 한글 버전이 설치된 환경에서 특히 문제
- **해결**: 관리자 권한으로 실행:
  ```
  "C:\Program Files (x86)\Hnc\Office 2024\HOffice130\Bin\Hwp.exe" /RegServer
  ```
- 또는 "한글 기본 설정 2024" → "기본값 복원"

### Breaking Change 6: 매크로 에러 (2025.06 업데이트)

- `HAction.Run("")`이 3번째 줄부터 실패
- COM 객체 재등록 불완전이 원인
- **해결**: 위와 동일 `/RegServer`

---

## 우리 코드에 필요한 대응

### 현재 OK인 것
- [x] ProgID `HWPFrame.HwpObject` 사용 → 전 버전 호환
- [x] 보안 모듈 `HwpAutomation` 경로 등록
- [x] `Quit()` 호출 (Drop trait에서)
- [x] out-of-process COM이라 32/64bit 무관

### 추가 고려 필요
- [ ] **Visible 토글 워크어라운드** — 2024 대응: set_visible에서 False→True 토글
- [ ] **`/RegServer` 안내** — 진단(diagnose)에 "COM 등록 상태 확인" 추가
- [ ] **다중 인스턴스 주의** — 2024에서 2개 이상 동시 생성 지양
- [ ] **버전 감지** — COM 객체의 `Version` 프로퍼티로 버전 확인 (10.x=2018, 11.x=2020, 12.x=2022, 13.x=2024)
- [ ] **CLSID 하드코딩 제거** — `find_install_path`에서 `{2291CF00-...}` 하드코딩 → ProgID에서 CLSID 동적 조회로 변경 (2018/2020은 `{2291CEFF-...}`)
- [ ] **다중 버전 설치 감지** — 여러 버전 설치 시 마지막 `/RegServer` 실행한 버전만 동작. 진단에 경고 추가
- [ ] **성능 팁 문서화** — Visible=False로 백그라운드 실행 시 10배 속도 향상, CreateAction 캐싱, PutFieldText 배치 입력

---

## 우리에게 유리한 점

1. **late-binding 사용** — Rust IDispatch 기반이라 .NET Interop RPC 에러 무관
2. **out-of-process COM** — 32/64bit 자동 마샬링, 비트 불일치 문제 없음
3. **ProgID 기반 연결** — `DispatchObject::create("HWPFrame.HwpObject")` → CLSID 차이 자동 해결
4. **Quit() 보장** — Drop trait에서 호출하므로 좀비 프로세스 방지

## 출처

- https://forum.developer.hancom.com/t/2024-hwpctrl/1013
- https://forum.developer.hancom.com/t/topic/1715
- https://forum.developer.hancom.com/t/2024-2/2235
- https://forum.developer.hancom.com/t/2024-registermodule/1655
- https://forum.developer.hancom.com/t/topic/2192
- https://forum.developer.hancom.com/t/250624-2024/2616
- https://forum.developer.hancom.com/t/c-hwp-2022-32bit/1894
- https://pyhwpx.com/67
- https://pyhwpx.com/30
- https://eduhub.co.kr/board_vRlj10/182153
