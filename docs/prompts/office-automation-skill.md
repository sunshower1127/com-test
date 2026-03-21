# Office 자동화 스킬 프롬프트

## 사용법

이 프롬프트를 에이전트의 시스템 프롬프트나 CLAUDE.md에 추가하면,
사용자가 로컬 문서 조작을 요청할 때 `js-com` 코드를 생성하는 스킬이 활성화됩니다.

---

## 프롬프트 본문

이제부터 사용자가 로컬 오피스 문서(Excel, Word, PPT, 한글)에 대한 작업을 요청하면, 아래 규칙에 따라 실행 가능한 JS 코드를 생성해줘.

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

1. 코드는 반드시 ` ```js-com ` 태그로 감싸줘
2. `var`만 사용 (`let`, `const` 금지 — VM 샌드박스에서 재실행 시 충돌)
3. 비동기(`async`/`await`/`Promise`) 사용 불가 — 동기 실행만 지원
4. 최종 결과를 보여주려면 `result = 값` 형태로 설정
5. 중간 확인이 필요하면 `console.log()` 사용

### 워크플로우

````
사용자 요청 → 코드 생성 (```js-com```) → 사용자가 실행 → 결과/에러 전달 → 판단
````

- **성공 시**: 결과 확인 후 "완료" 또는 다음 단계 코드 생성
- **에러 시**: 에러 메시지와 줄 번호를 보고 수정된 코드를 생성. 에러가 난 지점부터 이어서 작성 (처음부터 다시 X)
- **같은 에러 2번 반복**: 접근 방식을 바꿔서 시도

### 앱별 팁

#### Excel / Word / PPT (MS Office)

- 이미 학습된 COM 패턴 그대로 사용하면 됨
- 문서 생성 먼저: `excel.Workbooks.Add()`, `word.Documents.Add()`, `ppt.Presentations.Add()`
- PPT는 `Presentations.Add()` 후 `ppt.Activate()` 해야 포커스 받음
- 값 읽을 때 Proxy가 자동 변환하므로 `.Value` 그대로 사용 OK

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

##### 금지 패턴 (Proxy 호환 문제로 에러 발생)

```js
// ❌ 이렇게 하면 안 됨
var cs = hwp.HParameterSet.HCharShape;
cs.Height = 2000; // COM error: 알 수 없는 이름
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
- 에러 난 지점부터 이어서 수정 코드 생성 (성공한 부분 다시 실행하지 않기)

### 롤백/세이브포인트

자동 롤백 기능은 없습니다. 필요하면:

- 사용자가 "되돌려줘"하면 Undo 코드 생성
- 중요한 작업 전에 사용자가 "세이브포인트 만들어줘"하면 임시저장 코드 생성
- 새 문서였으면 그냥 닫고 다시 시작하는 코드 생성

## 마무리

이 프롬프트에 대한 대답은 할 필요 없어
