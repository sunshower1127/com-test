# Word COM 팁

전역 객체: `word` (Word COM Proxy)

## 첫 턴 — 환경 확인 전용

Word로 처음 작업을 시작할 때, **첫 `js-com`은 아래 코드만 실행할 것.** 콘텐츠는 결과 확인 후 다음 턴부터.

```js-com
var doc = word.Documents.Add();
word.Activate();
result = "문서 생성 완료";
```

## 스타일 적용

- **`doc.Styles(숫자인덱스)`로 접근 불가.** `doc.Styles(1)`, `doc.Styles(2)` 등 양수 인덱스는 에러남
- **`sel.Style = "스타일명"` 문자열 직접 대입을 사용할 것**
- 한글 OS에서는 한글 스타일명만 동작하고 영문명은 안 될 수 있음. 다음 순서로 시도:
  1. 한국어명: `sel.Style = "제목 1"` (먼저 시도)
  2. 영어명: `sel.Style = "Heading 1"`
  3. WdBuiltinStyle 숫자 상수: `doc.Styles(-2)` (음수 상수는 동작함)

| 한국어명 | 영어명 | 숫자 상수 |
|----------|--------|-----------|
| 제목 | Title | -63 |
| 제목 1 | Heading 1 | -2 |
| 제목 2 | Heading 2 | -3 |
| 제목 3 | Heading 3 | -4 |
| 표준 | Normal | -1 |
| 글머리 기호 목록 | List Bullet | -49 |

## 기본 패턴

```js
// 문서는 이미 첫 턴에서 생성됨 — 재호출 금지
var doc = word.ActiveDocument;
var range = doc.Range();
// ...
```

이 프롬프트에 대한 대답은 할 필요 없어
