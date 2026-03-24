# 단계 2: 텍스트 입출력

## 1. 공식 API 문서 조사

### 출처

- [한컴 개발자 - GetTextFile](https://developer.hancom.com/en-us/webhwp/devguide/hwpctrl/methods/gettextfile)
- [한컴 개발자 - SetTextFile](https://developer.hancom.com/webhwp/devguide/hwpctrl/methods/settextfile)
- [한컴 개발자 포럼 - GetTextFile Q&A](https://forum.developer.hancom.com/t/gettextfile/925)
- [pyhwpx - 텍스트 입력](https://wikidocs.net/257896)
- [pyhwpx - GetText with tables](https://pyhwpx.com/129)

### InsertText (HAction + HParameterSet)

```
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "텍스트"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

- 커서 위치에 텍스트 삽입
- `\r\n`으로 줄바꿈 가능
- 새 문단은 만들지 않음 — 문단 나누기는 `Run("BreakPara")` 별도 호출

### InsertText (CreateAction 패턴)

```
act = CreateAction("InsertText")
set = act.CreateSet()
act.GetDefault(set)
set.SetItem("Text", "텍스트")
act.Execute(set)
```

- HParameterSet 패턴과 동일한 결과

### GetTextFile(Format, option)

- **Format**: `"UNICODE"`, `"TEXT"`, `"HTML"`, `"HWPML2X"`, `"HWP"`
- **option**: `""` = 전체 문서, `"saveblock"` = 선택 영역만

### SetTextFile(data, Format, option)

- **data**: 삽입할 텍스트
- **Format**: `"UNICODE"`, `"TEXT"` 등
- **option**: `""` 또는 `"insertfile"`
- **공식 문서**: option `""` = 전체 교체, `"insertfile"` = 커서 위치에 삽입

### GetPageText(pgno, option)

- **pgno**: 0-based 페이지 번호
- 해당 페이지의 텍스트 반환

### GetHeadingString()

- 파라미터 없음
- 현재 커서 위치의 문단 제목/번호 반환 (예: "1.", "가)")
- 제목이 없으면 빈 문자열

### Run("BreakPara")

- 파라미터 없는 단순 액션
- 커서 위치에 문단 나누기 (Enter 키와 동일)

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 | 실험 결과 | 비고 |
|------|-----------|-----------|------|
| SetTextFile option="" | 전체 교체 | **커서 위치에 삽입** (교체 아님) | 공식 문서와 다름 |
| GetTextFile "UTF8" | 일부 문서에서 언급 | **미지원** (null 반환) | HwpObject에서 미지원 |
| GetTextFile "TEXT" | 텍스트 반환 | CP949/ANSI로 반환 (한글 깨짐) | Node.js에서 사용 불가 |
| saveblock + SelectText() | 선택 영역 추출 | **빈 문자열** 반환 | Run 기반 선택만 작동 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)
- 보안모듈 미설치 상태

### InsertText

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 2-1 | InsertText (HParameterSet) | ✅ | ret=true, 텍스트 정상 삽입 |
| 2-2 | InsertText (CreateAction) | ✅ | ret=true, 동일한 결과 |
| 2-3 | 여러 줄 삽입 `\r\n` | ✅ | 줄바꿈 정상 |

### GetTextFile

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 2-4 | GetTextFile("UNICODE", "") | ✅ | 정확한 유니코드 텍스트 반환 |
| 2-5 | GetTextFile("TEXT", "") | ⚠️ | 작동하지만 CP949 인코딩 → 한글 깨짐 |
| 2-6 | GetTextFile("UTF8", "") | ❌ | null 반환 — 미지원 |
| 2-7 | GetTextFile("HTML", "") | ✅ | 완전한 HTML 문서 (9KB) |
| 2-8 | GetTextFile("HWPML2X", "") | ✅ | 완전한 HWPML XML (28KB, UTF-16) |
| 2-9 | GetTextFile("HWP", "") | ✅ | Base64 인코딩 OLE 바이너리 (19KB) |

### SetTextFile

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 2-10 | SetTextFile("UNICODE", "") 전체 교체 | ⚠️ | **교체가 아닌 삽입** — 기존 내용 앞에 추가됨 |
| 2-11 | SetTextFile("TEXT", "") 전체 교체 | ⚠️ | 동일 — 교체 아닌 삽입 |
| 2-12 | SetTextFile("UNICODE", "insertfile") | ✅ | 커서 위치에 삽입 |
| 2-13 | SelectAll → Delete → SetTextFile | ✅ | **우회 방법으로 전체 교체 가능** |

### 기타

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 2-14 | GetPageText(0, "") | ✅ | 0-based, 첫 페이지 텍스트 정상 |
| 2-15 | GetHeadingString() | ✅ | 빈 문자열 (제목 없는 문단에서 정상) |
| 2-16 | Run("BreakPara") x2 | ✅ | 빈 줄 2개 정상 생성 |

### saveblock (선택 영역 추출)

| # | 선택 방법 | 결과 | 비고 |
|---|----------|:----:|------|
| 2-17 | SelectText(0,0,0,5) → saveblock | ❌ | 빈 문자열 (SelectionMode=1인데도) |
| 2-18 | MoveSelRight x5 → saveblock | ✅ | 정확히 5글자 반환 |
| 2-19 | Run("SelectAll") → saveblock | ✅ | 전체 텍스트 반환 |
| 2-20 | SelectText(0,0,0,10) → saveblock | ❌ | 빈 문자열 |

---

## 4. 분석

### 핵심 발견 1: InsertText vs SetTextFile

| | InsertText | SetTextFile |
|--|-----------|------------|
| 입력 | 순수 텍스트 | 포맷 문자열 (HTML, HWPML2X 등) |
| 서식 | ❌ 플레인 텍스트만 | ✅ 서식/표/이미지 포함 가능 |
| 용도 | 텍스트 타이핑 | 서식 있는 콘텐츠 로드 |

- InsertText에 `<b>볼드</b>` 넣으면 → 그대로 문자열로 보임
- SetTextFile에 `<b>볼드</b>` + format="HTML" → **볼드** 서식 적용

"File"은 파일 시스템이 아니라 **파일 포맷**을 의미.
디스크를 거치지 않는 `SaveAs` / `Open`의 인메모리 버전이다.
`GetTextFile("HWPML2X")` → 수정 → `SetTextFile(modified, "HWPML2X", "")` 로 서식 보존 라운드트립 가능.

### 핵심 발견 2: SetTextFile option="" 은 삽입이다

공식 문서는 `""` = 전체 교체라고 하지만, 실제로는 **커서 위치에 삽입**한다.
`"insertfile"`과 동작이 동일하다. HwpCtrl(ActiveX) 기준 문서와 HwpObject(OLE) 동작 차이로 추정.

**전체 교체 우회 방법:**
```js
hwp.Run("SelectAll");
hwp.Run("Delete");
hwp.SetTextFile(newContent, "UNICODE", "");
```

### 핵심 발견 2: SelectText()와 Run 기반 선택은 다르다

| | SelectText() API | Run("MoveSelRight") 등 |
|--|-----------------|----------------------|
| SelectionMode | 1 | 1 |
| saveblock | ❌ 빈 문자열 | ✅ 정상 |
| 성격 | 프로그래밍 선택 | UI 시뮬레이션 선택 |

`SelectText()`는 내부적으로 다른 종류의 선택 상태를 만든다.
**saveblock을 사용하려면 반드시 Run 기반 선택 명령을 사용해야 한다.**

### 핵심 발견 3: GetTextFile 포맷별 용도

| Format | 용도 | 비고 |
|--------|------|------|
| `"UNICODE"` | **기본 텍스트 추출** — 가장 실용적 | 항상 사용 |
| `"TEXT"` | ANSI/CP949 텍스트 | Node.js에서 한글 깨짐, 사용 금지 |
| `"UTF8"` | 미지원 | null 반환 |
| `"HTML"` | 서식 포함 HTML | 큰 결과물, 필요 시 사용 |
| `"HWPML2X"` | 완전한 문서 XML | 라운드트립 가능, 서식 보존 |
| `"HWP"` | Base64 OLE 바이너리 | 라운드트립 가능, 사람이 읽을 수 없음 |

### 핵심 발견 4: BreakPara가 자동교정(빠른 교정)을 트리거한다

HWP의 "빠른 교정(Quick Correct)" 기능은 1,374개 사전 기반 자동 치환 기능이다.
`Run("BreakPara")` 실행 시 **커서가 있는 문단만** 자동교정이 발동된다.

```
"첫번째" → BreakPara → "첫 번째"  (띄어쓰기 자동 삽입)
```

#### 트리거 조건

| 동작 | 자동교정 발동 |
|------|:------------:|
| `Run("BreakPara")` | ⚠️ 발동 |
| `InsertText`에 `\r\n` 포함 | ✅ 안 됨 |
| `MoveDocBegin/End` | ✅ 안 됨 |
| `MovePos` | ✅ 안 됨 |
| `MoveLeft/Right` | ✅ 안 됨 |
| `Cancel` | ✅ 안 됨 |

#### 범위

- **커서가 있는 문단만** 교정됨 (다른 문단은 영향 없음)
- 이미 교정된 문단에 다시 BreakPara해도 이중 교정 안 함

#### 비활성화 방법

| 방법 | 결과 |
|------|:----:|
| `Run("AutoSpellCheck")` | ❌ 효과 없음 |
| `Run("ToggleAutoCorrect")` | ❌ 효과 없음 |
| `HParameterSet.HQCorrect` 조작 | ❌ on/off 플래그 없음 |
| 레지스트리 변경 (실행 중) | ❌ 효과 없음 |
| **API로 끄는 방법은 존재하지 않음** | |

#### 우회: `\r\n` 사용

`InsertText`에 `\r\n`을 포함하면 BreakPara와 동일한 문단 구조(`<P>` 태그)를 만들면서 자동교정이 발동되지 않는다.

```js
// ❌ 자동교정 발동
insert('첫번째');
Run('BreakPara');
insert('두번째');

// ✅ 자동교정 안 됨, 동일한 문단 구조
insert('첫번째\r\n두번째');
```

**주의: 문서 중간에 커서를 놓고 `insert('\r\n')`하면 예상과 다르게 동작** (문단이 안 나뉘고 앞에 빈 줄 생김). 기존 문서 중간에 줄을 나눌 때는 BreakPara가 필요하며, 이 경우 자동교정은 감수해야 한다.

| 상황 | 권장 방법 |
|------|----------|
| 텍스트를 새로 쓸 때 | `insert('줄1\r\n줄2\r\n줄3')` |
| 기존 문서 중간에 줄 나누기 | `Run("BreakPara")` (자동교정 감수) |

---

## 5. 최종 결론

### 정상 작동 (✅)

- **InsertText** — 두 패턴 모두 정상 (HParameterSet, CreateAction)
- **GetTextFile("UNICODE")** — 텍스트 추출의 기본
- **GetTextFile("HTML" / "HWPML2X" / "HWP")** — 서식 포함 추출
- **SetTextFile("UNICODE", "insertfile")** — 커서 위치 삽입
- **GetPageText(0)** — 0-based 페이지별 텍스트
- **GetHeadingString()** — 제목 문자열
- **Run("BreakPara")** — 문단 나누기 (단, 자동교정 트리거 — 아래 주의 참고)
- **saveblock** — Run 기반 선택 후 사용 시 정상

### 주의 필요 (⚠️)

- **Run("BreakPara")** — 커서 문단에 자동교정 발동. 새 텍스트 작성 시 `\r\n`으로 대체 권장
- **SetTextFile option=""** — 전체 교체가 아닌 삽입. 교체하려면 SelectAll+Delete 선행
- **GetTextFile("TEXT")** — CP949 인코딩, Node.js에서 한글 깨짐
- **saveblock + SelectText()** — 빈 문자열 반환. Run 기반 선택만 사용할 것

### 미지원 (❌)

- **GetTextFile("UTF8")** — null 반환
