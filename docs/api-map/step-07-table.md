# 단계 7: 표 (Table)

## 1. 공식 API 문서 조사

### 표 생성: HAction + HParameterSet.HTableCreation

```js
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;  // 단에 맞춤
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
// → 커서가 자동으로 첫 셀 안으로 이동
```

### 셀 이동 (Run 액션)

| 액션 | 설명 |
|------|------|
| `TableRightCell` | 오른쪽 셀 (마지막 셀이면 다음 행 첫 셀) |
| `TableLeftCell` | 왼쪽 셀 |
| `TableUpperCell` | 위 셀 |
| `TableLowerCell` | 아래 셀 |

### HTableCreation 프로퍼티

| 프로퍼티 | 설명 |
|----------|------|
| Rows | 행 수 |
| Cols | 열 수 |
| WidthType | 0=사용자 지정, 2=단에 맞춤 |
| WidthValue | WidthType=0일 때 폭 (HWP 단위) |
| HeightType | 0=자동 |
| HeightValue | HeightType 지정 시 높이 |
| ColWidth | 열 폭 |
| RowHeight | 행 높이 |
| TableProperties | 표 속성 |
| TableTemplate | 표 템플릿 |

---

## 2. 공식 문서와 차이

| 항목 | 공식 문서 | 실제 | 비고 |
|------|-----------|------|------|
| 표 밖에서 재진입 | MoveDown/MoveRight 등 | **표 안으로 직접 진입 불가** | 표 생성 직후 연속 작업 권장 |
| CellShape | 셀 서식 접근 | 프로퍼티 존재하지만 Item 없음 | 개별 셀 서식은 CharShape/ParaShape로 적용 |

---

## 3. 테스트 결과

### 환경

- 한글 2018 (10.0.0.5060)

### 표 생성

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 7-1 | TableCreate 3x2 (HParameterSet) | ✅ | 커서 자동으로 첫 셀 진입 |
| 7-2 | TableCreate 2x3 (CreateAction) | ✅ | |

### 셀 이동

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 7-3 | TableRightCell | ✅ | |
| 7-4 | TableLowerCell | ✅ | |
| 7-5 | TableLeftCell | ✅ | |
| 7-6 | TableUpperCell | ✅ | |

### 셀 내 작업

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 7-7 | InsertText | ✅ | 셀 안에 텍스트 입력 |
| 7-8 | CharShape (볼드, 색상, 크기) | ✅ | 셀 내 서식 적용 정상 |

### 표 탈출/진입

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 7-9 | MoveDocEnd (표 밖으로) | ✅ | 표 뒤로 커서 이동 |
| 7-10 | MoveDocBegin → MoveRight (재진입) | ❌ | 표 안으로 안 들어감 (List=0, Para=1) |
| 7-11 | MoveNextPosEx (재진입) | ❌ | 동일 |

### GetTextFile

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 7-12 | GetTextFile("UNICODE") | ✅ | 셀 내용이 줄 단위로 추출됨 |

### CellShape

| # | 테스트 | 결과 | 비고 |
|---|--------|:----:|------|
| 7-13 | hwp.CellShape 접근 | ⚠️ | 객체는 존재하지만 Item/ItemExist 모두 비어있음 |

---

## 4. 분석

### 핵심 발견 1: 표 생성 → 셀 이동 → 텍스트/서식 전부 정상

표를 만들면 커서가 자동으로 첫 셀에 들어간다. 이후 `TableRightCell` 등으로 이동하면서 `InsertText`, `CharShape` 적용이 모두 정상 작동한다.

```js
// 표 생성
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
hwp.HParameterSet.HTableCreation.Rows = 3;
hwp.HParameterSet.HTableCreation.Cols = 2;
hwp.HParameterSet.HTableCreation.WidthType = 2;
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet);

// 셀 채우기
insert("A1");
hwp.Run("TableRightCell");
insert("B1");
hwp.Run("TableLowerCell");
insert("B2");
// ...
```

### 핵심 발견 2: 표 밖에서 재진입이 어려움

표를 나간 후 `MoveDocBegin`, `MoveRight`, `MoveNextPosEx` 등으로 표 안에 다시 들어갈 수 없다. 표 컨트롤은 커서가 건너뛴다.

**권장 패턴: 표 생성 직후 연속 작업**

```js
// ✅ 표 만들고 바로 채우기
hwp.HAction.Execute("TableCreate", ...);
insert("A1");
hwp.Run("TableRightCell");
insert("B1");
// ... 전부 채운 후
hwp.Run("MoveDocEnd");  // 표 밖으로 나가기
```

표에 다시 들어가야 하는 경우는 추가 조사 필요 (ctrl 접근, GetAnchorPos 등).

### 핵심 발견 3: GetTextFile은 표 내용을 포함

`GetTextFile("UNICODE")`는 표의 각 셀 내용을 줄 단위로 반환한다:

```
(빈줄 — 표 시작)
A1
B1
A2
B2
A3
B3
(빈줄 — 표 끝)
```

### 핵심 발견 4: CellShape은 사실상 미사용

`hwp.CellShape` 객체는 존재하지만 Item이 비어있어 사용 불가. 셀 내 서식은 `CharShape` / `ParaShape`로 적용.

---

## 5. 최종 결론

### 정상 작동 (✅)

- **TableCreate** — 표 생성 (HParameterSet, CreateAction 모두)
- **셀 이동** — TableRightCell, TableLeftCell, TableUpperCell, TableLowerCell
- **셀 내 InsertText** — 정상
- **셀 내 CharShape** — 볼드, 색상, 크기 등 정상
- **MoveDocEnd** — 표 밖으로 나가기
- **GetTextFile** — 표 내용 포함하여 추출

### 셀 배경색 (CellBorderFill)

```js
hwp.HAction.GetDefault("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
var hcbf = hwp.HParameterSet.HCellBorderFill;
var fa = hcbf.FillAttr;
fa.type = 1;
fa.WindowsBrush = 1;
fa.WinBrushFaceColor = hwp.RGBColor(255, 0, 0);  // 빨강
fa.WinBrushFaceStyle = -1;                         // ← 반드시 -1
fa.WinBrushHatchColor = hwp.RGBColor(0, 0, 0);    // ← 반드시 설정
hwp.HAction.Execute("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
```

**주의사항:**
- `WinBrushFaceStyle`: **-1**=단색, **생략**=단색(기본), **0~5**=빗금 패턴 (Windows HBRUSH 스타일: 0=가로줄, 1=세로줄, 2=대각선↘, 3=대각선↗, 4=격자, 5=X격자)
- `WinBrushHatchColor`는 선택사항 (없어도 됨)
- comPut, Proxy 접근 모두 작동

### 셀 테두리

```js
hwp.HAction.GetDefault("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
hwp.HParameterSet.HCellBorderFill.BorderTypeTop = 1;     // 실선
hwp.HParameterSet.HCellBorderFill.BorderWidthTop = 3;     // 굵기
hwp.HParameterSet.HCellBorderFill.BorderColorTop = hwp.RGBColor(255, 0, 0);
// Bottom, Left, Right도 동일하게
hwp.HAction.Execute("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet);
```

**참고:** `BorderCorlorLeft` — 오타 아님, HWP API가 `Corlor`로 되어있음

### 표 재진입 (표 밖에서 다시 안으로)

**방법 1: MoveDown으로 진입**
```js
hwp.Run("MoveDocBegin");
hwp.Run("MoveDown");          // 표 안으로 진입 (첫 열의 어딘가)
hwp.Run("TableUpperCell");    // 반복해서 첫 행으로
hwp.Run("TableLeftCell");     // 반복해서 첫 열로 → A1
```

**방법 2: SetPos(list, para, pos)로 직접 점프**
```js
// list ID를 알면 원하는 셀로 직접 이동
hwp.SetPos(4, 0, 0);  // list=4인 셀로
```

- 각 셀은 고유한 **List ID**를 가짐 (GetPosBySet().Item("List")로 확인)
- List ID는 문서마다 다르므로, 진입 후 읽어서 사용
- MoveDown은 첫 열의 첫 행이 아닌 다른 셀로 갈 수 있음 — TableUpperCell/LeftCell로 보정 필요

### 주의 필요 (⚠️)
- **CellShape** — 객체는 있지만 실질적으로 사용 불가. CharShape/ParaShape 사용
- **셀 배경색 읽기** — GetDefault가 기본값만 반환. 현재 셀 배경색을 읽는 방법 미확인
