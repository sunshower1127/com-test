# PPT, Word COM 확장 가능성 분석

## 상태: 🔲 미착수

## 목표

Excel과 동일한 구조(Proxy → napi-rs → Rust COM)로
PowerPoint, Word도 바로 지원 가능한지 확인.

## 확인 항목

- [ ] ProgID 확인: `PowerPoint.Application`, `Word.Application` 등록 여부
- [ ] `comCreate("PowerPoint.Application")` → Launch 되는지
- [ ] Proxy 패턴 호환성 — Excel처럼 직접 property get/put이 잘 되는지
- [ ] Word 특유의 패턴 (Selection, Range, Paragraph 모델)
- [ ] PPT 특유의 패턴 (Slides, Shapes, TextFrame 모델)
- [ ] 기존 com-core + napi-rs 코드 수정 없이 되는지, 앱별 bridge crate가 필요한지
- [ ] Electron worker에 launch/quit 명령 추가 범위

## 예상

Excel이 되면 Word/PPT도 거의 그대로 동작할 것.
COM의 IDispatch late-binding이라 ProgID만 바꾸면 될 가능성 높음.
