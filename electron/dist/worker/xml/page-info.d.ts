/**
 * 페이지 경계 감지 — COM으로 MovePageBegin/End + GetPosBySet
 */
import { ComBridge, ComHandle, PageBoundary } from './types';
/** 각 페이지의 시작/끝 위치 감지 */
export declare function detectPageBoundaries(bridge: ComBridge, hwpHandle: ComHandle): PageBoundary[];
