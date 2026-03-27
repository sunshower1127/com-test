# 유즈케이스: 양식 문서 자동 채우기 (Template Fill)

## 1. 문제 정의

사용자가 이력서 같은 양식 문서를 주고 "이 정보로 채워줘"라고 요청.
표가 포함된 HWP 문서에서 특정 셀에 데이터를 프로그래밍으로 넣어야 한다.

**핵심 난이도:**
- 이력서 등 양식은 복잡한 병합 셀로 구성
- COM MoveDown으로 표 진입은 가능하지만, **어느 셀에 도착하는지 예측 불가능**
- 서식(폰트, 배경, 테두리)을 100% 보존해야 함
- 이미지 삽입도 필요할 수 있음

---

## 2. 최종 파이프라인

테스트를 통해 확립된 안정적인 파이프라인:

```
1. COM으로 문서 열기 (hwp.Open)
2. HWPML2X로 문서 전체 XML 추출 (hwp.GetTextFile("HWPML2X", ""))
3. 코드로 XML 파싱:
   - HEAD(스타일) 제거 → BODY만 추출
   - TABLE → ROW → CELL 구조에서 구조맵 생성
   - 각 셀은 RowAddr/ColAddr로 식별, 텍스트 내용 포함
   - 88KB XML → ~1KB 구조맵으로 경량화
4. 구조맵(수 KB)을 LLM에게 전달 + 사용자 요청
5. LLM이 fill(table, row, col, value) 형태로 지시 반환
6. 코드가 원본 XML에서 해당 CELL 찾아 <CHAR>값</CHAR> 삽입
7. SetTextFile("HWPML2X")로 라운드트립 → 서식 100% 보존
```

### 구조맵 예시

```
## 표1 (이력서)
| "성명" | {빈칸:r0c1} | "생년월일" | {빈칸:r0c3} |
| "연락처" | {빈칸:r1c1} | "비상연락망" | {빈칸:r1c3} |
| "E-Mail" | {빈칸:r2c1} |
| "주소" | {빈칸:r3c1} |
| "학력사항" | "재학기간" | "학교명(전공)" |
| {빈칸:r5c0} | {빈칸:r5c1} | {빈칸:r5c2} |
```

### LLM 응답 형식

```json
[
  { "table": 0, "row": 0, "col": 1, "value": "홍길동" },
  { "table": 0, "row": 0, "col": 3, "value": "1990. 01. 15" },
  { "table": 0, "row": 1, "col": 1, "value": "010-1234-5678" }
]
```

**LLM은 COM을 직접 만지지 않고, 원본 XML도 보지 않는다.** 구조맵만 보고 지시를 내린다.

---

## 3. 이미지 삽입

HWPML2X 라운드트립으로는 **텍스트만** 삽입 가능하다. 이미지가 필요한 경우 별도 처리:

```
1. XML 편집 시 이미지가 필요한 셀에 플레이스홀더 삽입: {{PHOTO}}
2. SetTextFile("HWPML2X")로 라운드트립 (텍스트 + 플레이스홀더 반영)
3. Find("{{PHOTO}}")로 커서 이동
4. 플레이스홀더 텍스트 삭제 (선택 → Delete)
5. InsertPicture(path, ...)로 이미지 삽입
6. 이미지 크기 조절: 삽입 후 ctrl Properties로 Width/Height 설정
```

**주의:**
- InsertPicture는 항상 원본 크기로 삽입됨 (Width/Height 파라미터 무시)
- 삽입 후 별도로 크기 조절 필수
- 로컬 파일만 가능 (URL 불가)

---

## 4. 대안 비교

| 방법 | 장점 | 단점 |
|------|------|------|
| **HWPML2X 라운드트립** | 서식 100% 보존, 복잡한 표 OK | 문서 전체 왕복 |
| **COM 셀 이동 (MoveDown + TableRightCell)** | 직접적 | MoveDown 진입 셀 예측 불가, 탐색 필요 |
| **AllReplace** | 간단 | 플레이스홀더 필요, 1회성 |
| **HTML 파싱** | 구조 파악 가능 | 라운드트립 불가 |

**결론:** HWPML2X 라운드트립이 가장 안정적인 방법. COM 셀 이동도 가능하지만 MoveDown 진입 위치가 예측 불가능하여 탐색 로직이 필요하다.

### MoveDown 표 진입

MoveDown은 **어떤 표든 진입 가능**하다 (복잡한 이력서 표 포함). 단:
- 표 위에 텍스트가 있으면 여러 번 MoveDown 필요
- **어느 셀에 도착하는지 예측 불가능** (레이아웃 기반)
- List ID가 0보다 크면 표 안에 있는 것

```js
// 안전한 표 진입 패턴
hwp.Run("MoveDocBegin");
while (true) {
  hwp.Run("MoveDown");
  var list = hwp.GetPosBySet().Item("List");
  if (list > 0) break;  // 표 안으로 진입됨
}
// 이후 TableUpperCell/LeftCell로 원하는 셀 탐색
```

---

## 5. 대량 문서 대응

LLM은 전체 XML을 보지 않는다. 구조맵(수 KB)만 본다.

- **100페이지 문서**도 특정 표만 파싱 가능
- **단계적 표시** 전략:
  1. 문서 개요: 표 개수, 각 표의 크기/제목
  2. 특정 표 상세: 셀 구조 + 텍스트
  3. 특정 셀: 정확한 RowAddr/ColAddr + 현재 내용
- 88KB XML → HEAD 제거 시 ~37KB → 구조맵 추출 시 **~1KB**
- LLM context 부담 최소화

---

## 6. 병합 셀 처리

**옵션 B: null 채우기로 격자 유지**

병합된 셀이 있으면 구조맵에서 해당 위치를 null로 채워 격자를 유지한다.

```
| "성명" | null   | {빈칸} | null   |  ← ColAddr 0,1은 병합, 2,3은 병합
| "연락처" | null | {빈칸} | null   |
```

이렇게 하면:
- LLM이 row/col 좌표를 정확히 계산 가능
- 병합된 셀은 첫 번째 좌표로만 접근
- null 위치에 fill 지시하면 무시

---

## 7. Markdown to HWPX 변환 가능성

HWPX = ZIP + XML 구조이므로, 이론적으로 프로그래밍 생성 가능:

```
1. 헤더 템플릿 준비 (기본 HWPX의 header.xml, settings.xml 등)
2. Markdown AST 파싱
3. AST → section0.xml 변환 (단락, 표, 이미지 등)
4. ZIP으로 묶어 .hwpx 파일 생성
```

**장점:** HWP COM 없이 문서 생성 가능
**제한:** 복잡한 서식은 직접 XML 매핑 필요, 기존 양식 재활용 불가
