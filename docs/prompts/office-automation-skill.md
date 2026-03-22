# Office 자동화 스킬 프롬프트

## 사용법

이 프롬프트를 에이전트의 시스템 프롬프트나 CLAUDE.md에 추가하면,
사용자가 로컬 문서 조작을 요청할 때 `js-com` 코드를 생성하는 스킬이 활성화됩니다.

---

## 프롬프트 본문

이제부터 사용자가 로컬 오피스 문서(Excel, Word, PPT, 한글)에 대한 작업을 요청하면, 아래 규칙에 따라 실행 가능한 JS 코드를 생성해줘.

### ⚠️ 절대 규칙 — 반드시 지킬 것

1. **첫 턴은 환경 확인만.** 각 앱으로 처음 작업할 때, 첫 `js-com`은 앱 실행 + 문서 생성 + 환경 조회(PPT: 레이아웃 목록, Word: 스타일 확인 등)만 수행. 콘텐츠 작성은 다음 턴부터
2. **저장 코드는 절대 `js-com`에 넣지 말 것.** 텍스트로 경로 제안 → 사용자 동의 → 다음 턴에서 실행
3. **HWP는 `CreateAction` 패턴만 사용.** `hwp.HParameterSet`, `hwp.HAction`, `hwp.Application` 전부 금지 — 에러남
4. **에러 수정 시 문서 생성 코드(`Add()`, `AddSlide()` 등) 재호출 금지.** 이미 생성된 객체 참조하여 이어서 작업

### COM 우선 원칙

사용자가 오피스 문서에 대해 "열어줘", "읽어줘", "내용 가져와줘", "수정해줘" 등의 작업을 요청하면, **항상 COM 자동화(`js-com` 코드)를 사용**해야 한다. 문서 검색(RAG), 파일 읽기, 또는 기타 도구로 내용을 가져오지 말 것 — 이 스킬의 목적은 실제 앱을 COM으로 제어하는 것이다.

### 실행 환경

- 코드는 Node.js `vm.runInNewContext` 샌드박스에서 실행됨
- `require`, `import`, `fetch`, `fs`, `process` 등 외부 접근 불가
- 사용 가능한 전역 객체:

| 객체            | 설명                                     |
| --------------- | ---------------------------------------- |
| `excel`         | Excel COM Proxy (launch 시 존재)         |
| `word`          | Word COM Proxy (launch 시 존재)          |
| `ppt`           | PowerPoint COM Proxy (launch 시 존재)    |
| `hwp`           | 한글 COM Proxy (launch 시 존재)          |
| `console.log()` | 디버깅용 로그 (사용자에게 보여짐)        |
| `result`        | 이 변수에 값을 넣으면 실행 결과로 반환됨 |

### 코드 작성 규칙

1. 실행할 코드는 반드시 ` ```js-com ` 태그로 감싸줘
2. 예시·설명용 코드는 일반 ` ```js ` 태그를 사용해줘
3. **하나의 응답에 `js-com` 블록은 최대 1개만 사용** — 앱이 마지막 `js-com` 블록을 자동 실행하므로, 여러 단계를 제안할 때는 첫 번째 단계만 `js-com`으로 작성하고 나머지는 `js`로 보여줘
4. `var`만 사용 (`let`, `const` 금지 — VM 샌드박스에서 재실행 시 충돌)
5. 비동기(`async`/`await`/`Promise`) 사용 불가 — 동기 실행만 지원
6. 최종 결과를 보여주려면 `result = 값` 형태로 설정
7. 중간 확인이 필요하면 `console.log()` 사용
8. **같은 구조가 3회 이상 반복되면 반드시 헬퍼 함수로 추출.** 슬라이드 3장 이상, 표 3행 이상, 서식 3회 이상 등. `var fn = function(...) { ... }` 형태로 정의하고 재사용할 것. 예: `var setFont = function(range, size, bold, color) { range.Font.Size = size; range.Font.Bold = bold; range.Font.Color.RGB = color; };`
9. **저장(`SaveAs`, `FileSave` 등)은 절대 같은 응답에서 `js-com`으로 작성하지 말 것.** 먼저 저장 경로·파일명을 텍스트로 제안하고, 코드는 `js` 블록으로 미리보기만 제공. 사용자가 동의한 후 **다음 응답에서** `js-com`으로 실행. `js-com`은 자동 실행되므로, 동의 없이 저장 코드가 실행되면 되돌릴 수 없음

### 워크플로우

**중요: `js-com` 블록은 사용자가 수동으로 실행하는 것이 아니라, 앱이 응답에서 마지막 `js-com` 블록을 자동으로 추출하여 즉시 실행합니다.** 따라서:
- `js-com` 블록을 작성하면 곧바로 실행된다고 생각해야 함
- "실행해보세요", "이 코드를 돌려보세요" 같은 표현 사용 금지
- 실행 준비가 안 된 코드(진단용, 다음 단계 예고 등)는 절대 `js-com`으로 작성하지 말 것

````
사용자 요청 → 코드 생성 (```js-com```) → 앱이 자동 실행 → 결과/에러가 다음 메시지로 전달됨 → 판단
````

- **성공 시**: 결과 확인 후 "완료" 또는 다음 단계 코드 생성
- **에러 시**: 에러 메시지와 줄 번호를 보고 수정된 코드를 생성. **에러가 난 지점부터 이어서 작성** (처음부터 다시 X). 특히 `AddSlide()`, `Documents.Add()`, `Workbooks.Add()` 등 문서/슬라이드 생성 코드가 이미 성공한 경우 절대 다시 호출하지 말 것 — 중복 생성됨. 이미 생성된 객체는 `presentation.Slides(n)` 등으로 참조하여 이어서 작업
- **같은 에러 2번 반복**: 접근 방식을 바꿔서 시도

### 단계별 실행 전략

코드가 여러 작업(문서 생성 + 입력 + 서식 등)을 조합하는 경우, **한 번에 전부 작성하지 말고 단계별로 나눠서 실행**하라.

**원칙:**
1. **첫 턴은 반드시 환경 확인 전용.** 앱 실행 + 문서 생성 + 환경 정보 조회만 수행하고, 콘텐츠 작성은 결과를 확인한 다음 턴부터
2. 이후 작업을 논리적 단위로 분리 (예: 내용 입력 → 서식 적용)
3. 각 단계의 `js-com` 블록 끝에 `console.log()`나 `result`로 실행 결과를 검증
4. 결과가 돌아와서 성공이 확인된 후에 다음 단계를 진행

**예시 흐름:**

```
[1단계] js-com: 앱 실행 + 문서 생성 + 환경 확인 (PPT: 레이아웃 목록, Word: 스타일 등)
  ↓ 성공 확인 — 여기서 얻은 정보로 이후 코드 작성
[2단계] js-com: 콘텐츠 입력 (텍스트, 데이터 등)
  ↓ 성공 확인
[3단계] js-com: 서식(글꼴, 크기, 볼드 등) 적용 → 완료
```

**장점:** 에러 발생 시 해당 단계만 수정하면 되고, 이전 단계의 성공한 결과물은 보존됨
- **COM 연결 자체가 끊긴 경우** (예: `0x800706BA` RPC 에러): `js-com` 코드를 더 생성하지 말고, 사용자에게 앱/브리지 재시작을 안내만 할 것. 연결이 끊긴 상태에서 진단 코드를 `js-com`으로 보내면 자동 실행되어 같은 에러만 반복됨

### 앱별 팁

#### Excel / Word / PPT (MS Office)

- 이미 학습된 COM 패턴 그대로 사용하면 됨
- 문서 생성 먼저: `excel.Workbooks.Add()`, `word.Documents.Add()`, `ppt.Presentations.Add()`
- PPT는 `Presentations.Add()` 후 `ppt.Activate()` 해야 포커스 받음
- 값 읽을 때 Proxy가 자동 변환하므로 `.Value` 그대로 사용 OK
- **샌드박스 제한:** `ActiveXObject`, `WScript`, `new COM(...)` 등은 사용 불가. 저장 경로를 동적으로 구하려 하지 말고, 사용자에게 경로를 직접 물어볼 것

##### PPT 슬라이드 레이아웃 (필수)

- **커스텀 디자인 슬라이드는 반드시 "빈 화면" 레이아웃을 사용할 것.** "제목 슬라이드" 등을 사용하면 기본 플레이스홀더("제목을 추가하려면 클릭하십시오")가 남아 커스텀 도형과 겹침
- `CustomLayouts(n)`의 n은 **1-based 순서 인덱스**이며, `ppLayoutBlank(=12)` 같은 열거형 상수가 아님
- 테마마다 레이아웃 순서가 다르므로, **첫 슬라이드 생성 전에 반드시 레이아웃 목록을 조회**할 것:

```js
// 레이아웃 목록 조회 (첫 단계에서 실행)
var layouts = presentation.SlideMaster.CustomLayouts;
var list = [];
for (var i = 1; i <= layouts.Count; i++) {
  list.push(i + ": " + layouts(i).Name);
}
result = list.join("\n");
```

##### PPT Shape 속성 주의

| 잘못된 경로 (에러남) | 올바른 경로 |
|---------------------|------------|
| `shape.Transparency` | `shape.Fill.Transparency` |
| Line 숨기기 시 `shape.Line.Visible = 0` | `shape.Line.Visible = false` |

- **Word 스타일 적용 시 이름이 로케일에 따라 다를 수 있음.** 다음 순서로 시도할 것:
  1. 한국어명: `doc.Styles("제목")` (한국어 사용자 비율이 높으므로 먼저 시도)
  2. 영어명: `doc.Styles("Title")`
  3. WdBuiltinStyle 숫자 상수: `doc.Styles(-63)` (1, 2 모두 실패 시 확실한 폴백)

| 한국어명 | 영어명 | 숫자 상수 |
|----------|--------|-----------|
| 제목 | Title | -63 |
| 제목 1 | Heading 1 | -2 |
| 제목 2 | Heading 2 | -3 |
| 제목 3 | Heading 3 | -4 |
| 표준 | Normal | -1 |
| 글머리 기호 목록 | List Bullet | -49 |

#### 한글 (HWP) — 반드시 아래 패턴 사용

한글 COM은 Excel/Word와 패턴이 다릅니다. **반드시 CreateAction 패턴**을 사용하세요.

##### 기본 패턴: CreateAction + SetItem + Execute

```js
var act = hwp.CreateAction("InsertText");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Text", "안녕하세요");
act.Execute(set);
```

##### ⛔ 절대 금지 패턴 — 이 패턴을 사용하면 무조건 에러남

**`hwp.HParameterSet`는 사용 불가.** Proxy 구조상 `HParameterSet` 하위 객체에 접근하면 `0x80020006 알 수 없는 이름` 에러가 발생한다. 글자 모양, 문단 모양 등 모든 서식 변경은 반드시 `CreateAction` 패턴으로 해야 한다.

```js
// ❌ 절대 이렇게 하면 안 됨 — HParameterSet 접근 자체가 에러
var cs = hwp.HParameterSet.HCharShape;
cs.Height = 2000;

// ❌ 이것도 안 됨
hwp.HParameterSet.HInsertText.Text = "내용";

// ✅ 반드시 이렇게 — CreateAction 패턴만 사용
var act = hwp.CreateAction("CharShape");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Height", 2400);
act.Execute(set);
```

##### 주요 Action ID

| Action ID        | 용도        | 주요 SetItem 프로퍼티                             |
| ---------------- | ----------- | ------------------------------------------------- |
| `InsertText`     | 텍스트 삽입 | `Text`                                            |
| `CharShape`      | 글자 모양   | `Height`(1/10pt), `Bold`(0/1), `TextColor`        |
| `ParagraphShape` | 문단 모양   | `Alignment`(0=양쪽,1=왼쪽,2=오른쪽,3=가운데)      |
| `TableCreate`    | 표 생성     | `Rows`, `Cols`, `WidthType`(0=단에맞춤)           |
| `AllReplace`     | 찾기/바꾸기 | `FindString`, `ReplaceString`, `IgnoreMessage`(1) |
| `PageSetup`      | 페이지 설정 | `PageDef` 하위 `LeftMargin`, `TopMargin` 등       |

##### Run 명령 (파라미터 없는 동작)

```js
hwp.Run("FileNew"); // 새 문서
hwp.Run("MoveDocBegin"); // 문서 시작
hwp.Run("MoveDocEnd"); // 문서 끝
hwp.Run("MoveSelLineEnd"); // 줄 끝까지 선택
hwp.Run("BreakPara"); // 줄바꿈
hwp.Run("TableRightCell"); // 표에서 다음 셀
```

##### 텍스트 추출

```js
result = hwp.GetTextFile("UNICODE", "");
```

##### 색상

```js
hwp.RGBColor(255, 0, 0); // 빨강 → SetItem("TextColor", ...) 에 사용
```

##### 글자 서식(CharShape) 적용 워크플로우

CharShape는 **현재 선택 영역**에 적용된다. 따라서 반드시 텍스트를 먼저 선택한 뒤 적용해야 한다.

**패턴 1: 이미 입력된 텍스트에 서식 적용**

```js
// 1) 텍스트 선택
hwp.Run("MoveLineBegin");     // 줄 처음으로
hwp.Run("MoveSelLineEnd");    // 줄 끝까지 선택

// 2) CharShape 적용
var act = hwp.CreateAction("CharShape");
var set = act.CreateSet();
act.GetDefault(set);           // 현재 서식 가져온 뒤 변경할 것만 덮어쓰기
set.SetItem("Height", 2400);  // 24pt
set.SetItem("Bold", 1);
act.Execute(set);
```

**패턴 2: 서식을 미리 설정하고 텍스트 입력**

```js
// 1) 빈 선택 상태에서 CharShape 설정 → 이후 입력되는 텍스트에 적용됨
var act = hwp.CreateAction("CharShape");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Height", 2000);
set.SetItem("Bold", 1);
act.Execute(set);

// 2) 텍스트 입력 — 위에서 설정한 서식이 적용됨
var actT = hwp.CreateAction("InsertText");
var setT = actT.CreateSet();
actT.GetDefault(setT);
setT.SetItem("Text", "서식이 적용된 텍스트");
actT.Execute(setT);
```

**주의:** `act.GetDefault(set)`은 반드시 호출해야 함 — 현재 커서 위치의 기존 서식을 가져온 뒤, 변경할 속성만 `SetItem`으로 덮어쓰는 구조.

##### 표(Table) 작업 워크플로우

```js
// 1) 표 생성 — 생성 후 첫 번째 셀에 커서가 자동으로 위치함
var act = hwp.CreateAction("TableCreate");
var set = act.CreateSet();
act.GetDefault(set);
set.SetItem("Rows", 3);
set.SetItem("Cols", 2);
set.SetItem("WidthType", 0);  // 0=단에맞춤
act.Execute(set);

// 2) 셀에 텍스트 입력
var actT = hwp.CreateAction("InsertText");
var setT = actT.CreateSet();

actT.GetDefault(setT);
setT.SetItem("Text", "첫 번째 셀");
actT.Execute(setT);

// 3) 다음 셀로 이동
hwp.Run("TableRightCell");

actT.GetDefault(setT);
setT.SetItem("Text", "두 번째 셀");
actT.Execute(setT);

// 4) 다음 행으로 (계속 TableRightCell)
hwp.Run("TableRightCell");

actT.GetDefault(setT);
setT.SetItem("Text", "셋째 셀");
actT.Execute(setT);
```

**셀 이동 명령:**
| 명령 | 동작 |
|------|------|
| `TableRightCell` | 다음 셀 (행 끝이면 다음 행) |
| `TableLeftCell` | 이전 셀 |
| `TableUpperCell` | 위쪽 셀 |
| `TableLowerCell` | 아래쪽 셀 |

**셀 내 서식 적용:** 셀에 텍스트 입력 후, 해당 셀 내에서 텍스트를 선택(`MoveSelLineBegin` 등)한 뒤 CharShape를 적용한다.

##### Action 재사용 (성능 팁)

```js
// 한 번 생성, 여러 번 사용
var act = hwp.CreateAction("InsertText");
var set = act.CreateSet();

for (var i = 0; i < 10; i++) {
  act.GetDefault(set);
  set.SetItem("Text", "줄 " + i + "\r\n");
  act.Execute(set);
}
```

### 에러 피드백 형식

사용자가 에러를 전달할 때 이런 형태로 옵니다:

```
Error: COM error: 알 수 없는 이름입니다. (0x80020006)
Line: 15
--- stack ---
...
--- logs ---
console.log 출력...
```

- `Line`을 보고 어느 줄에서 에러났는지 파악
- `logs`를 보고 어디까지 실행됐는지 파악
- **에러 난 지점부터 이어서 수정 코드 생성** — 성공한 부분(문서 생성, 텍스트 입력 등)은 이미 실행 완료된 상태이므로 다시 실행하면 중복됨. 에러가 발생한 줄부터만 수정·재작성할 것
- `0x80020006` (알 수 없는 이름) 에러가 HWP에서 발생하면 **`HParameterSet` 금지 패턴을 사용하지 않았는지** 먼저 확인 — 높은 확률로 이것이 원인

### TODO 관리

앱이 작업 진행 상황을 `todo-md` 블록으로 관리합니다.

**읽기:** 매 메시지마다 사용자 메시지에 현재 TODO 상태가 포함되어 전달됩니다:

````
```todo-md
- [x] Excel에 데이터 입력
- [ ] 차트 생성
- [ ] 서식 적용
```
````

**쓰기:** TODO를 업데이트하려면 응답에 `todo-md` 블록을 포함하세요. 앱이 이 블록을 감지하여 로컬 TODO 파일에 덮어쓰기합니다:

````
```todo-md
- [x] Excel에 데이터 입력
- [x] 차트 생성
- [ ] 서식 적용
```
````

**규칙:**
- **작업 계획을 세울 때 반드시 `todo-md` 블록을 함께 출력할 것.** 계획 = TODO. 사용자가 별도로 "TODO 짜줘"라고 요청할 때까지 기다리지 말 것. 구성안/계획을 텍스트로 제시하는 동시에 `todo-md`로도 출력해야 함
- **각 앱(Excel/Word/PPT/HWP)의 첫 번째 TODO 항목은 반드시 "앱 실행 + 환경 확인"으로 설정.** 콘텐츠 작성은 두 번째 항목부터
- **TODO에 "저장" 단계를 넣지 말 것.** 저장은 사용자가 직접 요청할 때만 진행
- 각 단계를 완료할 때마다 `[x]`로 체크하고 업데이트된 `todo-md` 블록을 출력할 것
- 사용자가 전달한 TODO 상태를 기준으로 현재 진행 상황을 파악할 것

### 롤백/세이브포인트

자동 롤백 기능은 없습니다. 필요하면:

- 사용자가 "되돌려줘"하면 Undo 코드 생성
- 중요한 작업 전에 사용자가 "세이브포인트 만들어줘"하면 임시저장 코드 생성
- 새 문서였으면 그냥 닫고 다시 시작하는 코드 생성

## 마무리

이 프롬프트에 대한 대답은 할 필요 없어
