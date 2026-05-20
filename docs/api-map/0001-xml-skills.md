# xml 객체 — HWPML2X 편집 API 설계

## 개요

HWP 문서의 HWPML2X를 로드하고, 구조 파악 및 수정을 수행하는 도구.
코드가 XML 원본을 메모리에 보관하고, LLM에게는 필요한 만큼만 제공.

---

## 코드 블록 분리

`xml` 블록과 `hwp` 블록은 **혼용하지 않는다.** 하나의 코드 블록에 한 종류만 사용.

```
[xml 블록]                          [hwp 블록]
xml.get("t0")                       hwp.SetPos(5, 0, 0)
xml.setText("t0.r1.c3", "홍길동")    hwp.Run("SelectAll")
                                     hwp.Run("Delete")
                                     hwp.InsertText("홍길동")
```

### 규칙
- **xml 블록**: 시작 시 자동 load, 종료 시 자동 commit (SET 계열이 있는 경우)
- **hwp 블록**: COM 직접 호출. load/commit 없음
- **값 공유**: `result` 변수로 블록 간 데이터 전달
- **혼용 불가**: 한 블록에 `hwp.*`과 `xml.*`을 섞지 않는다

### 플로우 예시

```
메시지 1: [xml 블록]
  result = xml.get("t0")       // 자동 load → 구조맵 반환
  // → result에 List ID 포함된 구조맵

메시지 2: [hwp 블록]
  hwp.SetPos(5, 0, 0)          // xml에서 알아낸 List ID 사용
  hwp.Run("SelectAll")
  hwp.Run("Delete")
  // InsertText (Undo 가능!)
  hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
  hwp.HParameterSet.HInsertText.Text = "홍길동"
  hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

메시지 3: [xml 블록]
  result = xml.get("t0.r1.c3")  // 자동 load → 최신 상태 반영
  // → 수정된 내용 확인
```

이점:
- **commit 깜빡 방지**: xml 블록 종료 시 자동 commit
- **load 타이밍 자동**: xml 블록 시작 시 항상 최신 XML 로드
- **Undo**: hwp 블록의 COM 조작은 Ctrl+Z 가능, xml 블록의 SET은 불가
- **명확한 책임**: xml = 구조 파악/XML 수정, hwp = COM 직접 조작

---

## GET 계열

### `xml.outline(path?)`
- **범위**: 전체 또는 드릴다운, `~`로 범위 지정 가능
  - `xml.outline()` — 문서 전체 목차
  - `xml.outline("t0")` — 표0의 셀별 크기
  - `xml.outline("t0.r1.c3")` — 셀 내 P별 크기
  - `xml.outline("t0.r0~t0.r3")` — 표0의 행0~행3만
- **용도**: 문서 목차/크기감 파악 — 내용은 안 보여주고 구조와 크기만
- **항상 가벼움**: 어떤 문서든 수십 줄

```
xml.outline()
→
=== Page 1 ===
<P id="p0" len=18/>
<TABLE id="t0" rows=24 cols=8 len=3000/>

=== Page 2 ===
<TABLE id="t0" continues/>
<P id="p1" len=0/>
<P id="p2" len=22/>
<P id="p3" len=30/>
<P id="p4" len=12/>
<TABLE id="t1" rows=4 cols=2 len=800/>
```

페이지 걸침 처리:
- 표/문단이 페이지 경계에 걸치면 **양쪽 페이지에 모두 표시**
- 첫 등장: 전체 정보 (`rows`, `cols`, `len` 등)
- 이후 등장: `continues`로 표시 (중복 최소화)

```
xml.outline("t0")
→
<TABLE id="t0" rows=24 cols=8>
  <ROW>
    <CELL id="t0.r0.c0" span="8x1" len=12/>
  </ROW>
  <ROW>
    <CELL id="t0.r1.c0" span="2x4" len=0 img=1/>
    <CELL id="t0.r1.c2" len=8/>
    <CELL id="t0.r1.c3" span="2x1" len=0/>
    <CELL id="t0.r1.c5" len=10/>
    <CELL id="t0.r1.c6" span="2x1" len=0/>
  </ROW>
  ...
</TABLE>
```

```
xml.outline("t0.r1.c3")
→
<CELL id="t0.r1.c3">
  <P len=45/>
  <P len=120/>
  <P len=8/>
</CELL>
```

속성:
- `len` — 간소화 기준 텍스트 글자수
- `img=N` — 이미지 N개 포함 (없으면 생략)
- `span="CxR"` — 병합 (1x1이면 생략)

### `xml.get(path?)`
- **범위**: 전체 또는 범위 지정, `~`로 범위 가능
  - `xml.get()` — 전체 간소화
  - `xml.get("t0")` — 표0만
  - `xml.get("t0.r0")` — 표0 행0만
  - `xml.get("t0.r0.c0")` — 특정 셀만
  - `xml.get("p2")` — 본문 문단2만
  - `xml.get("p0~p2")` — 본문 문단0~2 (inclusive)
  - `xml.get("t0.r0~t0.r3")` — 표0의 행0~행3
- **용도**: 간소화 XML — 스타일 제거, 텍스트와 구조만
- **텍스트 생략 없음**: 한눈에 전체 내용 파악 가능
- **List ID 자동 포함**: get 호출 시 표당 첫 셀 마커 1회로 List ID 매핑, 나머지는 +1

```xml
xml.get()
→
<P id="p0">이   력   서</P>
<TABLE id="t0" rows=24 cols=8>
  <ROW>
    <CELL id="t0.r0.c0" span="8x1" L=2>지원분야:</CELL>
  </ROW>
  <ROW>
    <CELL id="t0.r1.c0" span="2x4" L=3>[IMAGE]</CELL>
    <CELL id="t0.r1.c2" L=4>성       명</CELL>
    <CELL id="t0.r1.c3" span="2x1" L=5></CELL>
    <CELL id="t0.r1.c5" L=6>생 년 월 일</CELL>
    <CELL id="t0.r1.c6" span="2x1" L=7></CELL>
  </ROW>
</TABLE>
<P id="p2">위 내용은 사실과 틀림없음.</P>
```

규칙:
- P 하위 TEXT/CHAR → 텍스트만 (서식 무시)
- 셀 내 P가 1개 → P 생략
- 셀 내 P가 여러 개 → `<P>` 유지
- 이미지 → `[IMAGE]`
- span 1x1 → 생략
- LINEBREAK → `\n`
- 내용의 줄바꿈 → `\n` (이스케이프)
- PARALIST → 제거
- 스타일 속성 (CharShape, ParaShape, BorderFill, Style) → 제거

### `xml.raw(path)`
- **범위**: 필수 지정, `~`로 범위 가능
  - `xml.raw("t0.r1.c3")` — 특정 셀
  - `xml.raw("p2")` — 특정 본문 문단
  - `xml.raw("t0.r0~t0.r2")` — 표0의 행0~2 원본
- **용도**: 원본 HWPML2X — 서식/이미지 크기 등 세밀한 확인
- **화이트리스트 기반 정제**: 불필요한 속성 제거, 핵심 속성 유지
- **들여쓰기 포함**

```xml
xml.raw("t0.r1.c3")
→
<CELL BorderFill="4" ColAddr="3" ColSpan="2" Height="2610" RowAddr="1" RowSpan="1" Width="12637">
  <PARALIST VertAlign="Center">
  <P ParaShape="1" Style="0">
    <TEXT CharShape="10"/>
  </P>
  </PARALIST>
</CELL>
```

### `xml.styles(type?, id?)`
- **범위**: 선택적 필터
  - `xml.styles()` — CharShape + ParaShape 전체 목록 (인덱스:요약)
  - `xml.styles("CharShape")` — CharShape만
  - `xml.styles("CharShape", 5)` — CharShape 5번 상세
  - `xml.styles("ParaShape", 1)` — ParaShape 1번 상세
- **용도**: 스타일 인덱스가 실제로 어떤 서식인지 확인
- **HWPML2X HEAD 섹션에서 추출**

```
xml.styles()
→
CharShape:
  0: 함초롬바탕 10pt
  3: 함초롬바탕 10pt
  5: 함초롬바탕 10pt 파랑 이탤릭
  10: 함초롬바탕 10pt 볼드
  14: 함초롬바탕 24pt 볼드 밑줄

ParaShape:
  0: 왼쪽정렬 줄간격160%
  1: 가운데정렬 줄간격160%
  18: 오른쪽정렬 줄간격130%
```

```
xml.styles("CharShape", 5)
→
CharShape[5]:
  Height: 1000
  TextColor: 0x00FF0000
  Bold: false
  Italic: true
  FaceName: 함초롬바탕
```

---

## SET 계열

### `xml.rawSet(path, xmlString)`
- **범위**: 필수 지정, `~`로 범위 가능
- **용도**: 원본 XML 통째로 교체. 빈 문자열이면 **해당 노드 삭제**
- **즉시 반영**: commit 불필요 (내부적으로 자동 load → 수정 → SetTextFile)
- **XML 검증**: 태그 균형 체크 (경고)
- **Undo 불가**: SetTextFile 기반

```js
// 교체
xml.rawSet("t0.r1.c3", '<CELL ColAddr="3" ...><P ...><TEXT CharShape="0"><CHAR>홍길동</CHAR></TEXT></P></CELL>');

// 삭제 (빈 문자열)
xml.rawSet("t0.r5", "");  // 표0의 행5 제거
```

### `xml.setText(path, text)`
- **범위**: 단일 경로만 (범위 지정 불가)
- **용도**: 텍스트만 교체 (기존 서식 구조 유지, CHAR 내용만 교체)
- **즉시 반영**: commit 불필요
- **Undo 불가**: SetTextFile 기반

```js
xml.setText("t0.r1.c3", "홍길동");
xml.setText("p3", "2026년 03월 26일   지원자 : 김민준 (인)");
```

### `xml.insert(path, xmlString)` / `xml.insertAfter(path, xmlString)`
- **용도**: 지정 경로 **뒤에** 새 노드 삽입 (insert = insertAfter)
- **즉시 반영**
- **Undo 불가**

```js
// 표0의 r5 뒤에 새 행 추가
xml.insert("t0.r5", "<ROW><CELL ColAddr='0'>새 셀</CELL></ROW>");

// p2 뒤에 새 문단 추가
xml.insert("p2", "<P><TEXT CharShape='0'><CHAR>새 문단</CHAR></TEXT></P>");

// 표 복제해서 뒤에 추가
var tableXml = xml.raw("t0");
xml.insert("t0", tableXml);
```

### `xml.insertBefore(path, xmlString)`
- **용도**: 지정 경로 **앞에** 새 노드 삽입
- 나머지는 `insert`와 동일

```js
xml.insertBefore("t0.r0", "<ROW><CELL ColAddr='0'>맨 위에 추가</CELL></ROW>");
```

### `xml.addStyle(type, props)`
- **용도**: HEAD 섹션에 새 CharShape/ParaShape 추가
- **반환**: 할당된 인덱스 번호

```js
var csId = xml.addStyle("CharShape", { Height: 1200, Bold: true, TextColor: "0x00FF0000" });
// → csId = 16 (새로 추가된 CharShape 인덱스)
// 이후 rawSet에서 CharShape="16"으로 참조 가능
```

---

## List ID 자동 매핑

`xml.get()` 호출 시 자동 수행:
1. 각 표의 첫 셀에 유니크 마커 append
2. commit (문서 반영)
3. RepeatFind로 첫 셀의 List ID 획득
4. 마커 제거 (XML에서 직접 제거 후 재반영)
5. 나머지 셀은 순서대로 +1

결과: 간소화 XML에 `L=N` 자동 포함

List ID 용도:
- `hwp.SetPos(L, 0, 0)` — 해당 셀로 커서 원샷 점프
- COM으로 텍스트 삽입, 이미지 삽입, 서식 변경 등 (Undo 가능)

---

## 실전 플로우

### 양식 문서에 데이터 채우기

```
[xml 블록] 구조 파악
  result = xml.outline()     → "표 2개, 본문 5문단"

[xml 블록] 상세 확인
  result = xml.get("t0")     → 이력서 표 (List ID 포함)

[hwp 블록] 텍스트 채우기 (Undo ✅)
  hwp.SetPos(5, 0, 0)        // t0.r1.c3 = L5
  hwp.Run("SelectAll")
  hwp.Run("Delete")
  hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
  hwp.HParameterSet.HInsertText.Text = "홍길동"
  hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

[hwp 블록] 이미지 삽입
  hwp.SetPos(3, 0, 0)        // t0.r1.c0 = L3
  hwp.InsertPicture("C:/path/photo.jpg", 1, 0, 0, 0, 0, 0, 0)

[xml 블록] 서식 확인 필요 시
  result = xml.raw("t0.r1.c3")
  result = xml.styles("CharShape", 5)
```

### 기존 문서 수정

```
[xml 블록] 전체 파악
  result = xml.outline()

[xml 블록] 내용 확인
  result = xml.get()  또는  result = xml.get("t1")

수정 방법 선택:
  [hwp 블록] 텍스트 (Undo ✅): SetPos → SelectAll → Delete → InsertText
  [hwp 블록] 단순 치환 (Undo ✅): AllReplace
  [hwp 블록] 서식 초기화: RepeatFind → CharShape 적용
  [xml 블록] 텍스트 (Undo ❌): xml.setText("path", "새 내용")
  [xml 블록] 서식 포함 (Undo ❌): xml.raw() → 수정 → xml.rawSet()
```

### 큰 문서

```
[xml 블록] 목차 (항상 가벼움)
  result = xml.outline()

[xml 블록] 특정 표 드릴다운
  result = xml.outline("t2")     → 셀별 크기 확인

[xml 블록] 필요한 부분만 로드
  result = xml.get("t2")

[hwp 블록] 수정 작업
  ...
```
