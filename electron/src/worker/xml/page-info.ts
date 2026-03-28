/**
 * 페이지 경계 감지 — COM으로 MovePageBegin/End + GetPosBySet
 */

import { ComBridge, ComHandle, PageBoundary } from './types';

/** 각 페이지의 시작/끝 위치 감지 */
export function detectPageBoundaries(
  bridge: ComBridge,
  hwpHandle: ComHandle,
): PageBoundary[] {
  const boundaries: PageBoundary[] = [];

  // 현재 커서 저장
  const savedPos = bridge.comCallWith(hwpHandle, 'GetPosBySet', []);

  // 페이지 수 확인
  const pageCount = bridge.comGet(hwpHandle, 'PageCount') as number;
  if (!pageCount || pageCount <= 0) return boundaries;

  bridge.comCallWith(hwpHandle, 'Run', ['MoveDocBegin']);

  for (let p = 0; p < pageCount; p++) {
    // 페이지 시작
    bridge.comCallWith(hwpHandle, 'Run', ['MovePageBegin']);
    const startPs = bridge.comCallWith(hwpHandle, 'GetPosBySet', []);
    const startList = bridge.comCallWith(startPs, 'Item', ['List']) as number || 0;
    const startPara = bridge.comCallWith(startPs, 'Item', ['Para']) as number || 0;

    // 페이지 끝
    bridge.comCallWith(hwpHandle, 'Run', ['MovePageEnd']);
    const endPs = bridge.comCallWith(hwpHandle, 'GetPosBySet', []);
    const endList = bridge.comCallWith(endPs, 'Item', ['List']) as number || 0;
    const endPara = bridge.comCallWith(endPs, 'Item', ['Para']) as number || 0;

    boundaries.push({
      page: p + 1,
      startList,
      startPara,
      endList,
      endPara,
    });

    // 다음 페이지로
    bridge.comCallWith(hwpHandle, 'Run', ['MoveRight']);
  }

  // 커서 복원
  try {
    bridge.comCallWith(hwpHandle, 'SetPosBySet', [savedPos]);
  } catch (_) {
    // 복원 실패 시 무시
  }

  return boundaries;
}
