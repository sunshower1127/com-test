/**
 * 페이지 경계 감지 — COM으로 MovePageBegin/End + GetPosBySet
 * TODO: Phase 5에서 구현
 */
import { ComBridge, ComHandle } from './types';
export interface PageBoundary {
    page: number;
    startList: number;
    startPara: number;
    endList: number;
    endPara: number;
}
/** 각 페이지의 시작/끝 위치 감지 */
export declare function detectPageBoundaries(_bridge: ComBridge, _hwpHandle: ComHandle): PageBoundary[];
