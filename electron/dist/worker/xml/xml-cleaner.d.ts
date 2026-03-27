/**
 * XML 정제 — 화이트리스트 기반 속성 필터링, 간결/상세 포맷팅
 * 순수 함수: XML 입력 → 문자열 출력
 */
/** 본문 문단 → 간결 한 줄 */
export declare function compactParagraph(pXml: string, pIdx: number): string;
/** 표 → 간결 XML (스타일 제거, 텍스트만) */
export declare function compactTable(tableXml: string, tIdx: number, listIdMap: Record<string, number>): string;
/** 표 → 상세 XML (화이트리스트 속성 유지) */
export declare function cleanTable(tableXml: string, tableIdx: number, baseIndent: number, listIdMap: Record<string, number>): string;
/** 문단 → 상세 XML */
export declare function cleanParagraph(pXml: string, id: string | undefined, baseIndent: number): string;
/** 화이트리스트 기반 속성 필터 */
export declare function filterAttrs(tagName: string, attrStr: string): string;
