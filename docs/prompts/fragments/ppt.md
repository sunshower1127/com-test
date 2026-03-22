# PPT COM 팁

전역 객체: `ppt` (PowerPoint COM Proxy)

## 첫 턴 — 환경 확인 전용

PPT로 처음 작업을 시작할 때, **첫 `js-com`은 아래 코드만 실행할 것.** 콘텐츠는 결과 확인 후 다음 턴부터.

```js-com
var prs = ppt.Presentations.Add();
ppt.Activate();
var layouts = prs.SlideMaster.CustomLayouts;
var list = [];
for (var i = 1; i <= layouts.Count; i++) {
  list.push(i + ": " + layouts(i).Name);
}
result = "슬라이드 크기: " + prs.PageSetup.SlideWidth + "x" + prs.PageSetup.SlideHeight + "\n레이아웃:\n" + list.join("\n");
```

이 결과에서 "빈 화면" 또는 "Blank"의 인덱스 번호를 확인하고, 이후 슬라이드 추가 시 해당 인덱스를 사용할 것.

## 슬라이드 레이아웃 (필수)

- **커스텀 디자인 슬라이드는 반드시 "빈 화면" 레이아웃을 사용할 것.** "제목 슬라이드" 등을 사용하면 기본 플레이스홀더("제목을 추가하려면 클릭하십시오")가 남아 커스텀 도형과 겹침
- `CustomLayouts(n)`의 n은 **1-based 순서 인덱스**이며, `ppLayoutBlank(=12)` 같은 열거형 상수가 아님
- 테마마다 레이아웃 순서가 다르므로, **첫 턴에서 레이아웃 목록을 조회한 결과를 기반으로 "빈 화면"의 인덱스를 사용**할 것

## 기본 패턴

```js
// 프레젠테이션은 이미 첫 턴에서 생성됨 — 재호출 금지
// 레이아웃 인덱스는 첫 턴 조회 결과 참고
var blankIdx = 7; // 예시 — 실제 값은 조회 결과에 따라 다름
var slide = prs.Slides.AddSlide(prs.Slides.Count + 1, prs.SlideMaster.CustomLayouts(blankIdx));
```

## Shape 속성 주의

| 잘못된 경로 (에러남) | 올바른 경로 |
|---------------------|------------|
| `shape.Transparency` | `shape.Fill.Transparency` |
| `shape.Line.Visible = 0` | `shape.Line.Visible = false` |

## 슬라이드 크기 (참고)

와이드스크린 16:9: `SlideWidth = 720, SlideHeight = 405` (포인트 단위)

## 포커스

`ppt.Presentations.Add()` 후 `ppt.Activate()` 해야 포커스 받음

이 프롬프트에 대한 대답은 할 필요 없어
