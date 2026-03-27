# 유즈케이스 2: 기존 문서 수정 (Modify Existing Document)

## 1. 기본 플로우

```
사용자: "이 문서에서 표 수정해줘" (기존 HWP 파일 제공)
→ 문서 열기 → 구조 파악 → 수정 위치 특정 → 수정 실행
```

생성과 다른 점: **이전 코드가 없다.** 문서의 현재 상태가 유일한 진실.

---

## 2. 핵심 과제: 위치 특정

사용자 요청은 대부분 모호하다:
- "그 표 부분 수정해줘"
- "3페이지 쯤에 있는 거"
- "실험 결과 테이블 업데이트"

→ [위치 특정 가이드](location-finding-guide.md) 참고

---

## 3. 수정 플로우

### Phase 1: 문서 열기 + 전체 파악

```
1. hwp.Open(path)
2. GetTextFile("HWPML2X") → XML (코드 메모리에 보관)
3. 파서로 Level 0 구조 개요 생성:
   "본문 5문단, 표 3개, 이미지 2개
    표0: 24행 8열 (성명/생년월일/...)
    표1: 5행 4열 (학력사항)
    표2: 4행 2열 (성장과정/...)"
4. 개요를 LLM에게 전달
```

### Phase 2: 대상 특정

```
LLM: "표0의 상세 구조를 보여주세요"
→ 파서가 표0만 Level 1 구조맵 생성 → LLM에게 전달

또는:
사용자가 셀을 드래그 → saveblock으로 해당 영역 텍스트/XML 읽기
→ LLM에게 전달
```

### Phase 3: 수정 실행

```
수정 직전에 XML을 다시 가져옴 (사용자 수정 반영):
  GetTextFile("HWPML2X") → XML v_latest

수정 타입에 따라:
  텍스트 치환 → AllReplace
  셀 값 변경 → XML v_latest에서 수정 → SetTextFile
  이미지 → 플레이스홀더 + Find + InsertPicture
  서식 → COM 직접 (CharShape 등)
```

---

## 4. 사용자 동시 수정 문제

기존 문서 수정에서도 동일한 문제 발생:

```
1. LLM이 XML 가져옴 (v1)
2. LLM이 구조 파악 중...
3. 사용자가 직접 문서 수정 (v1 → v2)
4. LLM이 v1 기반으로 수정 시도
   → v1을 SetTextFile하면 사용자 수정(v2) 사라짐!
```

### 해결: "늦은 읽기" 패턴

```
구조 파악: GetTextFile → v1 (캐시)
  ... 시간 경과 ...
수정 실행 직전: GetTextFile → v_latest (최신)
  v_latest에서 수정 → SetTextFile
```

### 캐시 유효성 검사

```js
// 방법 1: IsModified 플래그
if (hwp.IsModified) {
  // 다시 가져오기
}

// 방법 2: 텍스트 해시 비교
var currentHash = hash(hwp.GetTextFile("UNICODE", ""));
if (currentHash !== cachedHash) {
  // 구조가 바뀌었을 수 있음 → XML + 구조맵 다시 생성
}
```

---

## 5. 대량 문서 전략

### 100페이지 논문 수정 시

```
Phase 1 (빠르게):
  GetTextFile("UNICODE") → 전체 텍스트 (수 KB)
  또는 GetPageText(n) → 페이지별 텍스트
  → 검색/RAG로 대상 페이지 특정

Phase 2 (정밀하게):
  GetTextFile("HWPML2X") → 전체 XML
  파서가 대상 표만 추출 → LLM에게 전달

Phase 3 (수정):
  XML에서 해당 부분만 수정 → SetTextFile
```

### RAG 연동

```
사용자: "실험 결과 업데이트해줘"
→ RAG 검색: "실험 결과" → 3페이지에 있음
→ GetPageText(2)로 확인
→ HWPML2X에서 3페이지 부근 표 찾기
→ 수정
```

---

## 6. 수정 타입별 최적 경로

| 수정 타입 | 최적 방법 | 비고 |
|-----------|----------|------|
| 텍스트 치환 (A→B) | AllReplace | 가장 빠르고 안전 |
| 빈 셀 채우기 | HWPML2X 라운드트립 | 구조맵 기반 |
| 중간에 문단 추가 | HWPML2X에 `<P>` 노드 삽입 | 미검증 |
| 행 추가/삭제 | HWPML2X에서 ROW 노드 조작 | 미검증 |
| 이미지 교체/추가 | Find + InsertPicture | 플레이스홀더 패턴 |
| 서식 변경 | COM 직접 (CharShape) | 위치 특정 필요 |
| 표 배경색 | COM CellBorderFill | WinBrushFaceStyle=-1 필수 |

---

## 7. 요약

```
기존 문서 수정의 핵심 원칙:

1. 문서의 현재 상태가 유일한 진실 (이전 코드 없음)
2. 구조 파악은 HWPML2X + 파서 (LLM에겐 경량 구조맵만)
3. 수정 직전에 항상 최신 XML을 다시 가져오기
4. 텍스트 치환은 AllReplace, 셀 수정은 HWPML2X 라운드트립
5. 사용자 동시 수정은 "늦은 읽기" 패턴으로 해결
```
