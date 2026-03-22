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

이름이 로케일에 따라 다를 수 있음. 다음 순서로 시도할 것:
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

## 기본 패턴

```js
// 문서는 이미 첫 턴에서 생성됨 — 재호출 금지
var doc = word.ActiveDocument;
var range = doc.Range();
// ...
```

이 프롬프트에 대한 대답은 할 필요 없어
