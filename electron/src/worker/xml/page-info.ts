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
export function detectPageBoundaries(
  _bridge: ComBridge,
  _hwpHandle: ComHandle,
): PageBoundary[] {
  // TODO: Phase 5에서 구현
  // 1. MoveDocBegin
  // 2. Loop: MovePageEnd + GetPosBySet → 페이지 경계 기록
  // 3. 원래 커서 복원
  return [];
}
