# LLM 프롬프트 설계

## 상태: 🔲 미착수

## 목표

LLM에게 오피스 자동화 코드를 생성시킬 때 사용할
시스템 프롬프트 + cheat sheet 주입 방식 설계.

## 확인 항목

- [ ] 시스템 프롬프트 구조 설계
  - 역할 정의 (오피스 자동화 코드 생성기)
  - 사용 가능한 객체 (`excel`, `hwp`, `console.log`, `result`)
  - 금지 사항 (`require`, `import`, `fetch` 등 sandbox 밖 접근 불가)
  - 에러 시 행동 가이드
- [ ] HWP cheat sheet 내용 확정
  - CreateAction 패턴 필수 명시
  - 주요 Action ID + SetItem 프로퍼티 목록
  - `Run()` 명령 목록
  - 반드시 피해야 할 패턴 (HParameterSet 직접 property put)
- [ ] Excel은 cheat sheet 필요한가?
  - LLM 학습 데이터로 충분한지 테스트
  - 필요하면 최소한의 힌트만 (Workbooks.Add 먼저 등)
- [ ] cheat sheet 주입 방식
  - 시스템 프롬프트에 직접 포함 vs 별도 메시지로 전달
  - 토큰 예산 고려 (HWP cheat sheet 크기)
- [ ] 에러 피드백 포맷
  - 에러 메시지 + 스택트레이스 + 원본 코드 + 줄 번호
  - console.log 출력 포함 여부
  - 동일 에러/코드 반복 감지 로직
- [ ] 재시도 정책
  - 최대 3회 (연구 기반)
  - 동일 에러 2회 연속 시 중단
  - 토큰 예산 상한

## 참고

- [how-modern-llms-debug-their-own-code-automatically.md](../references/how-modern-llms-debug-their-own-code-automatically.md)
- [proxy-debugging-notes.md](../analysis/proxy-debugging-notes.md) 의 LLM 가이드라인 섹션
- [hwp-com-api-reference.md](../analysis/hwp-com-api-reference.md)
