# xml 명령어 출력 형식 가이드

HWP 문서를 HWPML2X 기준으로 읽고 수정할 때 쓰는 `xml` 객체 명령어들의 **반환값 형식**과 **예시**를 정리한 문서입니다.

모든 GET/SET 명령은 **문자열**을 반환합니다. 코드 블록에서는 보통 `result` 변수에 담아 확인합니다.

```js
result = xml.outline();
result = xml.get("t0");
```

---

## 공통: 경로(path) 표기법

문서 안의 각 요소는 **id 경로**로 지정합니다.

| 표기          | 의미                              | 예시                   |
| ------------- | --------------------------------- | ---------------------- |
| `pN`          | 본문 N번째 블록 (문단 또는 표)    | `p0`, `p3`             |
| `tN`          | N번째 top-level 표                | `t0`, `t1`             |
| `tN.rR`       | 표 N, 행 R                        | `t0.r1`                |
| `tN.rR.cC`    | 표 N, 행 R, 열 C (`ColAddr` 기준) | `t0.r1.c3`             |
| `tN.rR.cC.pP` | 셀 안 P번째 문단                  | `t0.r5.c0.p2`          |
| `시작~끝`     | 범위 (양쪽 포함)                  | `p0~p2`, `t0.r0~t0.r3` |

**주의**

- `pN` 인덱스는 HWP COM의 Para 인덱스와 같습니다. **표도 P 하나를 차지**하므로 `p1`이 표일 수 있습니다.
- 셀 열 번호는 XML의 `ColAddr` 값입니다. 병합 셀은 `span="열x행"`으로 표시됩니다.
- 빈 본문 문단은 `outline()`에서 생략될 수 있습니다.

---

## GET 계열 — 읽기

읽기 명령은 xml 블록 시작 시 **자동 `load()`** 됩니다. 반환값은 항상 **하나의 문자열**입니다.

### `xml.outline(path?)` — 목차/크기만

**용도:** 문서 구조와 크기 파악. **텍스트 내용은 보여주지 않음.**  
**특징:** 어떤 문서든 출력이 짧음 (수십 줄 수준).

#### `xml.outline()` — 문서 전체 목차

각 top-level 요소(본문 P, 표)를 한 줄 요약합니다.

```
<P id="p0" len=9/>
<TABLE id="t0" rows=24 cols=8 len=3821/>
<P id="p3" len=78/>
<P id="p4" len=32/>
<P id="p5" len=9/>
<TABLE id="t1" rows=4 cols=2 len=1054/>
```

| 속성           | 의미                                                 |
| -------------- | ---------------------------------------------------- |
| `id`           | 경로 id (`p0`, `t0` …)                               |
| `len`          | 텍스트 글자 수 (CHAR 기준)                           |
| `rows`, `cols` | 표 행/열 수                                          |
| `len` (TABLE)  | 해당 표를 `get()`으로 간소화했을 때 대략적인 문자 수 |

**페이지 정보가 감지되면** 페이지별로 묶이고, 페이지를 걸치는 요소는 `continues`로 표시됩니다.

```
=== Page 1 ===
<P id="p0" len=18/>
<TABLE id="t0" rows=24 cols=8 len=3000/>

=== Page 2 ===
<TABLE id="t0" continues/>
<P id="p3" len=22/>
```

#### `xml.outline("t0")` — 표 내부 셀별 크기

```
<TABLE id="t0" rows=24 cols=8>
  <ROW>
    <CELL id="t0.r0.c0" span="8x1" len=87/>
  </ROW>
  <ROW>
    <CELL id="t0.r1.c0" span="2x4" len=0/>
    <CELL id="t0.r1.c2" len=10/>
    <CELL id="t0.r1.c3" span="2x1" len=0/>
    <CELL id="t0.r1.c5" len=8/>
    <CELL id="t0.r1.c6" span="2x1" len=0/>
  </ROW>
  ...
</TABLE>
```

| 속성         | 의미                            |
| ------------ | ------------------------------- |
| `span="CxR"` | 셀 병합 (열 x 행). 1x1이면 생략 |
| `img=N`      | 이미지 N개 포함 (없으면 생략)   |
| `len=0`      | 텍스트 없음 (이미지 전용 셀 등) |

#### `xml.outline("t0.r5.c0")` — 셀 내 문단별 크기

```
<CELL id="t0.r5.c0">
  <P id="t0.r5.c0.p0" len=1/>
  <P id="t0.r5.c0.p1" len=1/>
  <P id="t0.r5.c0.p2" len=1/>
  <P id="t0.r5.c0.p3" len=1/>
</CELL>
```

#### `xml.outline("t0.r0~t0.r2")` — 행 범위

```
<ROW>
  <CELL id="t0.r0.c0" span="8x1" len=87/>
</ROW>
<ROW>
  <CELL id="t0.r1.c0" span="2x4" len=0/>
  <CELL id="t0.r1.c2" len=10/>
  ...
</ROW>
<ROW>
  <CELL id="t0.r2.c2" len=5/>
  ...
</ROW>
```

---

### `xml.get(path?)` — 간소화된 내용 (구조 + 텍스트)

**용도:** 실제 텍스트와 구조를 함께 확인.  
**특징:** 서식 속성은 제거하고, **텍스트는 생략하지 않음**.

#### 출력 규칙

- `TEXT`/`CHAR` → 텍스트만 추출 (CharShape, ParaShape 등 제거)
- 셀에 P가 **1개** → `<P>` 태그 없이 셀 안에 바로 텍스트
- 셀에 P가 **여러 개** → `<P id="...">` 유지
- 이미지 → `[IMAGE]`
- 병합 셀 → `span="8x1"` 등
- 줄바꿈(LINEBREAK) → `\n` (리터럴 백슬래시+n 두 글자)
- `L=N` → COM `SetPos(N, 0, 0)`에 쓸 List ID (표 셀만, `get()` 호출 시 자동 매핑)

#### `xml.get()` — 문서 전체

```
<P id="p0">이   력   서</P>
<TABLE id="t0" rows=24 cols=8>
  <ROW>
    <CELL id="t0.r0.c0" span="8x1">                                                                       지원분야:           </CELL>
  </ROW>
  <ROW>
    <CELL id="t0.r1.c0" span="2x4"></CELL>
    <CELL id="t0.r1.c2" L=12> 성       명</CELL>
    <CELL id="t0.r1.c3" span="2x1" L=13></CELL>
    <CELL id="t0.r1.c5" L=14> 생 년 월 일</CELL>
    <CELL id="t0.r1.c6" span="2x1" L=15></CELL>
  </ROW>
  ...
</TABLE>
<P id="p3">                                                               위 내용은 사실과 틀림없음.</P>
```

> `L=` 값은 HWP가 열려 있을 때 COM으로 매핑한 결과입니다. 오프라인 XML 파일만으로는 붙지 않습니다.

#### `xml.get("p3")` — 특정 본문 문단

```
<P id="p3">                                                               위 내용은 사실과 틀림없음.</P>
```

#### `xml.get("t0.r1.c2")` — 특정 셀

```
<CELL id="t0.r1.c2" L=12> 성       명</CELL>
```

#### `xml.get("t0.r5.c0")` — 셀 내 여러 문단

```
<CELL id="t0.r5.c0" L=20>
  <P id="t0.r5.c0.p0">학</P>
  <P id="t0.r5.c0.p1">력</P>
  <P id="t0.r5.c0.p2">사</P>
  <P id="t0.r5.c0.p3">항</P>
</CELL>
```

#### `xml.get("p0~p1")` — 본문 범위

```
<P id="p0">이   력   서</P>
```

(`p1`이 표이면 P 타입만 포함되므로 표 내용은 나오지 않음)

#### `xml.get("t0.r0~t0.r2")` — 표 행 범위

```
  <ROW>
    <CELL id="t0.r0.c0" span="8x1" L=10>...</CELL>
  </ROW>
  <ROW>
    ...
  </ROW>
```

---

### `xml.raw(path)` — 원본 HWPML2X (정제본)

**용도:** CharShape, BorderFill, 크기 등 **서식·레이아웃 확인**  
**특징:** path **필수**. 불필요한 속성은 화이트리스트로 걸러진 **원본에 가까운 XML** 반환.

#### `xml.raw("p0")`

```xml
<P ColumnBreak="false" PageBreak="false" ParaShape="1" Style="0">
<TEXT CharShape="14">
<SECDEF CharGrid="0" FirstBorder="false" ...>
...
<CHAR>이   력   서</CHAR>
</TEXT>
</P>
```

#### `xml.raw("t0.r1.c2")`

```xml
<CELL BorderFill="23" ColAddr="2" ColSpan="1" ... Height="2610" ... Width="7574">
<CELLMARGIN Bottom="141" Left="141" Right="141" Top="141"/>
<PARALIST ... VertAlign="Center">
<P ParaShape="1" Style="0">
<TEXT CharShape="10">
<CHAR> 성       명</CHAR>
</TEXT>
<TEXT CharShape="3"/>
</P>
</PARALIST>
</CELL>
```

| raw vs get | raw                                      | get                |
| ---------- | ---------------------------------------- | ------------------ |
| 텍스트     | `<CHAR>` 태그 유지                       | 태그 없이 텍스트만 |
| 서식       | ParaShape, CharShape, BorderFill 등 유지 | 제거               |
| 크기       | Height, Width 등 유지                    | 제거               |
| 용도       | 서식 수정·복제                           | 내용 파악·List ID  |

---

### `xml.styles(type?, id?)` — HEAD 스타일 목록

**용도:** `CharShape="10"` 같은 숫자가 실제 어떤 서식인지 확인.

#### `xml.styles()` — 전체 요약

```
CharShape:
  0: 굴림 10pt 볼드
  1: 한양신명조 11pt
  2: 굴림 10pt
  ...
  10: 굴림 10pt 볼드
  14: 맑은 고딕 25pt 볼드 밑줄

ParaShape:
  0: Justify
  1: Center
  2: Left
  3: Right
  ...
```

#### `xml.styles("CharShape", 10)` — 특정 CharShape 상세

```
CharShape[10]:
  BorderFillId: 1
  Height: 1000
  Id: 10
  TextColor: 0
  UseFontSpace: false
  ...
  BOLD: /
```

#### `xml.styles("ParaShape", 1)` — 특정 ParaShape 상세

```
ParaShape[1]:
  Align: Center
  ...
  PARAMARGIN: Indent=0 Left=0 LineSpacing=160 LineSpacingType=Percent ...
  PARABORDER: BorderFill=1 Connect=false IgnoreMargin=false/
```

---

## SET 계열 — 쓰기

SET 명령도 **문자열**을 반환합니다. 내부 XML은 즉시 수정되고, xml 블록 종료 시 **자동 `commit()`** 으로 문서에 반영됩니다 (Undo 불가).

| 명령                                | 반환값 예시                   | 설명                              |
| ----------------------------------- | ----------------------------- | --------------------------------- |
| `xml.setText(path, text)`           | `setText t0.r1.c3 applied.`   | 텍스트만 교체, 서식 구조 유지     |
| `xml.rawSet(path, xmlString)`       | `SET t0.r1.c3 applied.`       | 노드 XML 통째 교체. `""`이면 삭제 |
| `xml.insert(path, xmlString)`       | `insert after t0.r5 applied.` | path **뒤에** 삽입                |
| `xml.insertAfter(path, xmlString)`  | 同上                          | `insert`와 동일                   |
| `xml.insertBefore(path, xmlString)` | `insertBefore t0.r0 applied.` | path **앞에** 삽입                |
| `xml.addStyle(type, props)`         | **숫자** (새 스타일 id)       | CharShape/ParaShape 추가          |

### 예시

```js
// 텍스트만 바꾸기
result = xml.setText("t0.r1.c3", "홍길동");
// → "setText t0.r1.c3 applied."

// 원본 XML로 교체
result = xml.rawSet(
  "t0.r1.c3",
  '<CELL ColAddr="3" ...><P><TEXT CharShape="0"><CHAR>홍길동</CHAR></TEXT></P></CELL>'
);

// 행 삭제
result = xml.rawSet("t0.r5", "");

// 새 CharShape 추가
var csId = xml.addStyle("CharShape", {
  Height: 1200,
  Bold: true,
  TextColor: "0x00FF0000",
});
// → 17 (할당된 인덱스)
```

---

## 유틸리티

| 명령               | 반환값                                    | 설명                         |
| ------------------ | ----------------------------------------- | ---------------------------- |
| `xml.load()`       | `XML loaded (123456 chars)`               | HWPML2X 수동 로드            |
| `xml.commit()`     | `Committed.` 또는 `No changes to commit.` | 변경사항 문서 반영           |
| `xml.mapListIds()` | `Mapped 48 cells.`                        | 모든 표 셀 List ID 일괄 매핑 |

xml 블록에서는 보통 `load`/`commit`을 직접 부르지 않아도 됩니다.

- 블록 **시작** → 자동 `load()`
- 블록 **종료** (SET이 있었으면) → 자동 `commit()`

---

## 명령어별 비교 요약

| 명령        | 출력 형식                     | 내용           | 크기       | List ID   |
| ----------- | ----------------------------- | -------------- | ---------- | --------- |
| `outline()` | XML-like 한 줄 요약           | ❌             | ✅ (`len`) | ❌        |
| `get()`     | 간소화 XML + 텍스트           | ✅             | 중간       | ✅ (`L=`) |
| `raw()`     | 정제된 HWPML2X                | ✅ (태그 포함) | 큼         | ❌        |
| `styles()`  | `인덱스: 요약` 또는 속성 목록 | 스타일만       | 작~중      | ❌        |

---

## 추천 워크플로

```
1. xml.outline()          → 문서가 큰지, 표/문단 몇 개인지 파악
2. xml.outline("t0")      → 필요한 표의 셀 크기 확인
3. xml.get("t0")          → 실제 텍스트 + L= List ID 확인
4. [hwp 블록] hwp.SetPos(L, 0, 0) …  → COM으로 수정 (Undo ✅)
   또는
   xml.setText("t0.r1.c3", "값")     → XML 직접 수정 (Undo ❌)
5. xml.raw("t0.r1.c3")    → 서식 문제 있을 때만 상세 확인
6. xml.styles("CharShape", 10)  → CharShape 번호 의미 확인
```

---

## 에러/특수 반환 메시지

경로가 잘못되면 명령별로 문자열 에러 메시지가 반환됩니다.

```
Table not found: t9 (표 9 없음, 최대 2개)
Cell not found: 지정한 경로의 셀을 찾을 수 없습니다
Paragraph not found: p20 (문단 20 없음, 최대 6개)
p5 위치는 TABLE입니다
Range not supported: ...
HEAD section not found
```
