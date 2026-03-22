# Excel COM 팁

전역 객체: `excel` (Excel COM Proxy)

## 첫 턴 — 환경 확인 전용

Excel로 처음 작업을 시작할 때, **첫 `js-com`은 아래 코드만 실행할 것.** 콘텐츠는 결과 확인 후 다음 턴부터.

```js-com
var wb = excel.Workbooks.Add();
excel.Activate();
result = "워크북 생성 완료. 시트: " + wb.Sheets(1).Name;
```

## 기본 패턴

```js
// 워크북은 이미 첫 턴에서 생성됨 — 재호출 금지
var ws = excel.ActiveWorkbook.Sheets(1);
ws.Cells(1, 1).Value = "데이터";
```

## 값 읽기

Proxy가 자동 변환하므로 `.Value` 그대로 사용 OK

## 범위 지정

```js
var rng = ws.Range("A1:D10");
rng.Font.Bold = true;
```

이 프롬프트에 대한 대답은 할 필요 없어
