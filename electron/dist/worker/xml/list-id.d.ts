/**
 * List ID 매핑 — COM 호출로 각 표 셀의 런타임 List ID를 알아내기
 * 현재: 모든 셀에 마커 삽입 → RepeatFind → 제거
 * TODO: 표당 첫 셀만 마커, 나머지는 +1 (간소화)
 */
import { ComBridge, ComHandle } from './types';
/** 모든 표 셀의 List ID를 매핑 */
export declare function mapListIds(xml: string, bridge: ComBridge, hwpHandle: ComHandle): Record<string, number>;
