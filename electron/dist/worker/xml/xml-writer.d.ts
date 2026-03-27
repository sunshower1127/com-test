/**
 * SET 계열 — XML 수정
 * 순수 함수: XML 입력 → 수정된 XML 출력 (COM 호출 없음)
 */
/** 특정 경로의 XML을 교체. 빈 문자열이면 삭제 */
export declare function rawSet(xml: string, path: string, newXml: string): string;
/** 텍스트만 교체 (기존 서식 구조 유지) */
export declare function setText(xml: string, path: string, value: string): string;
/** 지정 경로 뒤에 새 노드 삽입 */
export declare function insertAfter(xml: string, path: string, newXml: string): string;
/** 지정 경로 앞에 새 노드 삽입 */
export declare function insertBefore(xml: string, path: string, newXml: string): string;
/** HEAD 섹션에 새 CharShape/ParaShape 추가, 할당된 인덱스 반환 */
export declare function addStyle(xml: string, type: string, props: Record<string, unknown>): {
    xml: string;
    id: number;
};
