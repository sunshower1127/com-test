# COM API Capability Map — 마일스톤 계획

## 새로운 목표

COM을 통해 **되는 것 / 안 되는 것**을 체계적으로 파악하고,
검증된 작업을 **헬퍼 함수로 추상화**하여 프롬프트에서 직접 호출하도록 유도한다.

> 생 COM 코딩을 시키는 것보다, 검증된 함수 호출을 명시하는 게 훨씬 안정적이다.

## 원칙

- **기존 지식 백지화**: 이전에 "된다/안된다"고 판단한 내용을 전제하지 않는다. 모든 API를 처음부터 실험으로 검증한다.
- **실험 기반**: 문서나 추측이 아니라, 실제 호출 → 사용자 눈으로 확인한 결과만 신뢰한다.
- **점진적 축적**: 한 API씩 시도하고 기록하며, 결과가 쌓여야 다음 단계로 간다.

## 전체 흐름

```
[1] API 수집 → [2] 호출 시도 → [3] 사람이 결과 확인 → [4] 되는/안되는 분류
                                                            ↓
[5] 안 되는 것 원인 분석 → [6] 되는 것 헬퍼 함수화 → [7] 프롬프트에 함수 명시
```

## 앱 순서

| 순서 | 앱 | 이유 |
|------|----|------|
| 1 | 한글 (HWP) | COM 문서 부족, 지뢰밭 많음 → 가장 먼저 정리 필요 |
| 2 | Word | MS Office 중 가장 문서 조작에 가까움 |
| 3 | Excel | 데이터 조작 위주, 비교적 안정적 |
| 4 | PowerPoint | 슬라이드/도형 위주, 특성이 다름 |

---

## Milestone 1: HWP API Capability Map

### 1-1. API 목록 수집 (백지 상태에서)
- `list_members()`로 HWPFrame.HwpObject의 메서드/프로퍼티 전체 열거
- 하위 객체(CreateAction 반환값, XHwpWindows 등)도 재귀적으로 탐색
- 웹에서 한컴 개발자 문서 등 추가 API 확인
- **기존 cheat sheet나 프롬프트 내용은 참고하지 않음** — 실험 결과로만 판단
- **산출물**: `docs/api-map/hwp-raw-api.md` — 전체 메서드/프로퍼티 목록

### 1-2. API 호출 시도 (카테고리별)
각 API를 실제로 호출해보고, 사용자가 눈으로 결과 확인.

| 카테고리 | 예시 API | 확인 포인트 |
|----------|----------|-------------|
| 문서 생명주기 | Open, SaveAs, Clear, Quit, Run("FileNew") | 파일 열림/저장/닫힘 |
| 텍스트 입출력 | InsertText, GetTextFile | 텍스트 삽입됨/추출됨 |
| 커서 이동 | Run("MoveDocBegin"), Run("MoveDocEnd") 등 | 커서 위치 변경 |
| 서식 (글자) | CharShape: Height, Bold, Italic, TextColor | 서식 반영 여부 |
| 서식 (문단) | ParagraphShape: Alignment, LineSpacing | 서식 반영 여부 |
| 표 | TableCreate, TableRightCell 등 | 표 생성/탐색 |
| 찾기/바꾸기 | AllReplace | 치환 결과 |
| 페이지 설정 | PageSetup | 여백/용지 변경 |
| 그림/OLE | InsertPicture 등 | 삽입 여부 |
| 문서 정보 | 페이지 수, 커서 위치 조회 등 | 값 반환 |
| **문서 수정** | 기존 내용 선택 → 삭제/교체 | **핵심: 이게 되는지 안 되는지** |

- **산출물**: `docs/api-map/hwp-test-results.md` — 각 API별 결과 (✅/❌/⚠️)

### 1-3. 안 되는 것 원인 분석
- 에러 코드별 분류 (0x80020006, 0x80020003 등)
- Proxy 구조 한계 vs COM 자체 한계 vs 보안모듈 한계
- 워크어라운드 가능 여부
- **산출물**: `docs/api-map/hwp-limitations.md`

### 1-4. 헬퍼 함수 추상화
검증 완료된 작업을 JS 헬퍼 함수로 래핑.

```js
// 예시: hwpHelpers.js
var hwpHelpers = {
  insertTextWithStyle: function(hwp, text, fontSize, bold) { ... },
  createTable: function(hwp, rows, cols, data) { ... },
  replaceAll: function(hwp, find, replace) { ... },
  // ...
};
```

- 함수는 VM 샌드박스에 주입 가능한 형태
- **산출물**: `src/helpers/hwp-helpers.js`

### 1-5. 프롬프트 업데이트
- 기존 프롬프트에서 "이렇게 코딩해라" → "이 함수를 호출해라"로 전환
- 헬퍼 함수 시그니처 + 사용법을 프롬프트에 명시
- **산출물**: `docs/prompts/fragments/hwp.md` 업데이트

---

## Milestone 2: Word API Capability Map

### 2-1 ~ 2-5: HWP와 동일한 흐름
- `list_members()` → API 목록 수집
- 카테고리별 호출 시도 (Documents, Range, Selection, Find, Styles 등)
- 안 되는 것 분류 + 원인 분석
- 헬퍼 함수 추상화
- 프롬프트 업데이트

### Word 특이사항
- 백지 상태에서 실험. 기존에 "됐다"고 한 것도 다시 확인
- Range vs Selection, 스타일, Find/Replace 등 모두 직접 시도

---

## Milestone 3: Excel API Capability Map

### 3-1 ~ 3-5: 동일 흐름
- Workbooks, Worksheets, Range, Cells, Charts 등
- 수식/서식/차트 각각 테스트
- 헬퍼 함수화
- 프롬프트 업데이트

---

## Milestone 4: PowerPoint API Capability Map

### 4-1 ~ 4-5: 동일 흐름
- Presentations, Slides, Shapes, TextFrame 등
- 레이아웃, 도형, 텍스트 프레임 조작 테스트
- 헬퍼 함수화
- 프롬프트 업데이트

---

## Milestone 5: 통합 + 프롬프트 최종 구조

### 5-1. 헬퍼 함수 VM 주입
- 모든 앱의 헬퍼 함수를 VM 샌드박스에 자동 주입하는 구조
- `createComProxy` 시점에 헬퍼도 함께 로드

### 5-2. 프롬프트 최종 정리
- "되는 것만 모은" 최종 프롬프트
- 앱별 헬퍼 함수 카탈로그 (시그니처 + 예시)
- "안 되는 것" 경고 목록 (LLM이 시도하지 않도록)

### 5-3. 검증
- 대표 시나리오별 E2E 테스트
  - "보고서 작성해줘" (HWP)
  - "이 문서 수정해줘" (Word)
  - "데이터 정리해줘" (Excel)
  - "발표자료 만들어줘" (PPT)

---

## 작업 방식

각 앱의 1-2 단계(API 호출 시도)는 다음 프로세스를 따른다:

1. **내가 코드 생성** → com-cli 또는 js-com으로 API 호출
2. **사용자가 실행 + 결과 확인** → "됐다" / "안됐다" / "에러났다"
3. **결과 기록** → test-results 문서에 반영
4. **반복** → 다음 API로

> 사용자의 눈으로 확인하는 과정이 핵심. 자동화 테스트로는 "화면에 실제로 반영됐는지" 판단 불가.
