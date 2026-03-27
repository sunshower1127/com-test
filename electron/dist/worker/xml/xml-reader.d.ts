/**
 * GET 계열 — 문서 구조 읽기
 * 순수 함수: XML 입력 → 문자열 출력 (COM 호출 없음)
 */
/** 간결 구조맵 — 스타일 제거, 텍스트+구조만 */
export declare function structure(xml: string, listIdMap: Record<string, number>): string;
/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 */
export declare function structureDetail(xml: string, listIdMap: Record<string, number>): string;
