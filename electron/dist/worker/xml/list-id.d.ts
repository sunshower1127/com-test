/**
 * List ID 매핑 — COM 호출로 각 표 셀의 런타임 List ID를 알아내기
 *
 * 간소화 전략: 표당 첫 셀만 마커 삽입 → RepeatFind → 나머지는 +1
 */
import { ComBridge, ComHandle } from './types';
/** 지정 표들의 셀 List ID를 매핑 (간소화: 표당 첫 셀만 조회, 나머지 +1 추론) */
export declare function mapListIds(xml: string, bridge: ComBridge, hwpHandle: ComHandle, targetTables?: number[]): Record<string, number>;
