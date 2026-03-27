/**
 * GET 계열 — 문서 구조 읽기
 * 순수 함수: XML 입력 → 문자열 출력 (COM 호출 없음)
 */

import { extractSection, collectDocElements, parseTables } from './path-resolver';
import { compactParagraph, compactTable, cleanTable, cleanParagraph } from './xml-cleaner';

/** 간결 구조맵 — 스타일 제거, 텍스트+구조만 */
export function structure(xml: string, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  const out: string[] = [];
  let pIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      out.push(compactParagraph(el.content, pIdx));
      pIdx++;
    } else {
      out.push(compactTable(el.content, tIdx, listIdMap));
      tIdx++;
    }
  }

  return out.join('\n');
}

/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 */
export function structureDetail(xml: string, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  const out: string[] = [];
  let pIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      out.push(cleanParagraph(el.content, 'p' + pIdx, 0));
      pIdx++;
    } else {
      out.push(cleanTable(el.content, tIdx, 0, listIdMap));
      tIdx++;
    }
  }

  return out.join('\n');
}
