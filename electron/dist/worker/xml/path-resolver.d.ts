/**
 * 경로 파싱, XML 내 요소 위치 찾기
 * - parsePath("t0.r1.c3") → 구조화된 세그먼트
 * - findElement(xml, "t0.r1.c3") → { xml, start, end }
 * - collectDocElements(section) → 문서 순서대로 P/TABLE 수집
 */
import { CellInfo, DocElement, ElementLocation } from './types';
/** 태그에서 숫자 속성 추출 */
export declare function getAttr(tag: string, name: string): number | null;
/** XML 특수문자 이스케이프 */
export declare function escapeXml(str: string): string;
/** XML 엔티티 언이스케이프 */
export declare function unescapeXml(str: string): string;
/** SECTION 내용 추출 */
export declare function extractSection(xml: string): {
    section: string;
    sectionStart: number;
};
export interface PathSegment {
    type: 't' | 'r' | 'c' | 'p';
    index: number;
}
export interface ParsedPath {
    segments: PathSegment[];
    tableIdx?: number;
    rowIdx?: number;
    colIdx?: number;
    paraIdx?: number;
}
/** "t0.r1.c3.p0" → ParsedPath */
export declare function parsePath(path: string): ParsedPath;
/** "t0.r0~t0.r3" → { start, end } */
export declare function parseRange(path: string): {
    start: ParsedPath;
    end: ParsedPath | null;
};
/** 셀 XML + 위치 찾기 */
export declare function findCell(xml: string, tableIdx: number, rowIdx: number, colIdx: number): ElementLocation;
/** 테이블 XML + 위치 찾기 */
export declare function findTable(xml: string, tableIdx: number): ElementLocation;
/** 행 XML + 위치 찾기 (테이블 내) */
export declare function findRow(xml: string, tableIdx: number, rowIdx: number): ElementLocation;
/** 본문 문단 XML + 위치 찾기 (TABLE 밖) */
export declare function findParagraph(xml: string, paraIdx: number): ElementLocation;
/** 경로로 요소 찾기 (통합) */
export declare function findElement(xml: string, path: string): ElementLocation;
/** SECTION 내 P(본문)와 TABLE을 문서 순서대로 수집 */
export declare function collectDocElements(section: string): DocElement[];
/** 표를 CellInfo 3차원 배열로 파싱 */
export declare function parseTables(xml: string): CellInfo[][][];
/** CHAR 태그에서 텍스트 추출 */
export declare function extractText(cellXml: string): string;
