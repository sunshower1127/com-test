/** 공유 타입, 인터페이스, 상수 */
export type ComHandle = unknown;
export interface ComBridge {
    comGet(handle: ComHandle, prop: string): unknown;
    comPut(handle: ComHandle, prop: string, value: unknown): void;
    comCallWith(handle: ComHandle, method: string, args: unknown[]): unknown;
}
export interface CellInfo {
    col: number;
    row: number;
    colSpan: number;
    rowSpan: number;
    text: string;
}
export interface DocElement {
    type: 'p' | 'table';
    index: number;
    content: string;
}
export interface ElementLocation {
    xml: string;
    start: number;
    end: number;
}
export interface PageBoundary {
    page: number;
    startList: number;
    startPara: number;
    endList: number;
    endPara: number;
}
/** structureDetail용 화이트리스트 — 태그별 유지할 속성 */
export declare const ATTR_WHITELIST: Record<string, string[] | '*'>;
