/**
 * 경로 파싱, XML 내 요소 위치 찾기
 * - parsePath("t0.r1.c3") → 구조화된 세그먼트
 * - findElement(xml, "t0.r1.c3") → { xml, start, end }
 * - collectDocElements(section) → 문서 순서대로 P/TABLE 수집
 */

import { CellInfo, DocElement, ElementLocation } from './types';

// ──────── 유틸리티 ────────

/** 태그에서 숫자 속성 추출 */
export function getAttr(tag: string, name: string): number | null {
  const re = new RegExp(name + '="(\\d+)"');
  const m = re.exec(tag);
  return m ? +m[1] : null;
}

/** XML 특수문자 이스케이프 */
export function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** XML 엔티티 언이스케이프 */
export function unescapeXml(str: string): string {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"');
}

/** top-level TABLE 매칭 (중첩 테이블 안전 — depth 추적) */
export function matchTopLevelTables(xml: string): Array<{ match: string; index: number }> {
  const results: Array<{ match: string; index: number }> = [];
  let pos = 0;
  while (pos < xml.length) {
    const start = xml.indexOf('<TABLE', pos);
    if (start < 0) break;
    // depth 추적
    let depth = 0;
    let scan = start;
    while (scan < xml.length) {
      const nextOpen = xml.indexOf('<TABLE', scan + 1);
      const nextClose = xml.indexOf('</TABLE>', scan);
      if (nextClose < 0) { scan = xml.length; break; }
      // 첫 스캔 시 depth=0 → 1
      if (scan === start) { depth = 1; scan = start + 6; continue; }
      if (nextOpen >= 0 && nextOpen < nextClose) {
        depth++;
        scan = nextOpen + 6;
      } else {
        depth--;
        if (depth === 0) {
          const end = nextClose + 8; // '</TABLE>'.length
          results.push({ match: xml.substring(start, end), index: start });
          pos = end;
          break;
        }
        scan = nextClose + 8;
      }
    }
    if (depth !== 0) break; // 불완전한 TABLE
  }
  return results;
}

/** top-level TABLE을 마커로 치환 + 원본 매치 정보 반환 */
export function replaceTopLevelTables(xml: string, marker: string): { replaced: string; tables: Array<{ match: string; index: number }> } {
  const tables = matchTopLevelTables(xml);
  let replaced = xml;
  // 뒤에서부터 치환 (인덱스 안 밀리게)
  for (let i = tables.length - 1; i >= 0; i--) {
    const t = tables[i];
    replaced = replaced.substring(0, t.index) + marker + replaced.substring(t.index + t.match.length);
  }
  return { replaced, tables };
}

/** SECTION 내용 추출 */
export function extractSection(xml: string): { section: string; sectionStart: number } {
  const m = xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!m) throw new Error('SECTION not found');
  return {
    section: m[1],
    sectionStart: m.index! + m[0].indexOf(m[1]),
  };
}

// ──────── 경로 파싱 ────────

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
export function parsePath(path: string): ParsedPath {
  const segments: PathSegment[] = [];
  const parts = path.split('.');
  const result: ParsedPath = { segments };

  for (const part of parts) {
    const type = part[0] as 't' | 'r' | 'c' | 'p';
    const index = parseInt(part.substring(1));
    segments.push({ type, index });

    switch (type) {
      case 't': result.tableIdx = index; break;
      case 'r': result.rowIdx = index; break;
      case 'c': result.colIdx = index; break;
      case 'p': result.paraIdx = index; break;
    }
  }

  return result;
}

/** "t0.r0~t0.r3" → { start, end } */
export function parseRange(path: string): { start: ParsedPath; end: ParsedPath | null } {
  const parts = path.split('~');
  const start = parsePath(parts[0].trim());
  const end = parts.length > 1 ? parsePath(parts[1].trim()) : null;
  return { start, end };
}

// ──────── 요소 찾기 ────────

/** 셀 XML + 위치 찾기 */
export function findCell(xml: string, tableIdx: number, rowIdx: number, colIdx: number): ElementLocation {
  const tables = matchTopLevelTables(xml);
  let ti = 0;
  for (const tm of tables) {
    if (ti === tableIdx) {
      const tableStart = tm.index;
      const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(tm.match))) {
        if (ri === rowIdx) {
          const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
          let cm;
          while ((cm = cellRe.exec(rm[0]))) {
            const ca = getAttr(cm[0], 'ColAddr');
            if (ca === colIdx) {
              const absStart = tableStart + rm.index + cm.index;
              return { xml: cm[0], start: absStart, end: absStart + cm[0].length };
            }
          }
          throw new Error('Cell not found: t' + tableIdx + '.r' + rowIdx + '.c' + colIdx + ' (ColAddr=' + colIdx + ' 없음)');
        }
        ri++;
      }
      throw new Error('Row not found: t' + tableIdx + '.r' + rowIdx + ' (행 ' + rowIdx + ' 없음, 최대 ' + ri + '행)');
    }
    ti++;
  }
  throw new Error('Table not found: t' + tableIdx + ' (표 ' + tableIdx + ' 없음, 최대 ' + ti + '개)');
}

/** 테이블 XML + 위치 찾기 */
export function findTable(xml: string, tableIdx: number): ElementLocation {
  const tables = matchTopLevelTables(xml);
  if (tableIdx < tables.length) {
    const t = tables[tableIdx];
    return { xml: t.match, start: t.index, end: t.index + t.match.length };
  }
  throw new Error('Table not found: t' + tableIdx + ' (표 ' + tableIdx + ' 없음, 최대 ' + tables.length + '개)');
}

/** 행 XML + 위치 찾기 (테이블 내) */
export function findRow(xml: string, tableIdx: number, rowIdx: number): ElementLocation {
  const table = findTable(xml, tableIdx);
  const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
  let rm;
  let ri = 0;
  while ((rm = rowRe.exec(table.xml))) {
    if (ri === rowIdx) {
      const absStart = table.start + rm.index;
      return { xml: rm[0], start: absStart, end: absStart + rm[0].length };
    }
    ri++;
  }
  throw new Error('Row not found: t' + tableIdx + '.r' + rowIdx + ' (행 ' + rowIdx + ' 없음, 최대 ' + ri + '행)');
}

/** 본문 문단 XML + 위치 찾기 (paraIdx = COM Para 인덱스, P+TABLE 모두 카운트) */
export function findParagraph(xml: string, paraIdx: number): ElementLocation {
  const { section, sectionStart } = extractSection(xml);
  const elements = collectDocElements(section);

  let bodyIdx = 0;
  for (const el of elements) {
    if (bodyIdx === paraIdx) {
      if (el.type !== 'p') {
        throw new Error('p' + paraIdx + ' 위치는 TABLE입니다. t 경로를 사용하세요');
      }
      const absStart = sectionStart + el.index;
      return { xml: el.content, start: absStart, end: absStart + el.content.length };
    }
    bodyIdx++;
  }
  throw new Error('Paragraph not found: p' + paraIdx + ' (문단 ' + paraIdx + ' 없음, 최대 ' + bodyIdx + '개)');
}

/** 경로로 요소 찾기 (통합) */
export function findElement(xml: string, path: string): ElementLocation {
  const parsed = parsePath(path);
  const seg = parsed.segments;

  if (seg[0].type === 'p') {
    return findParagraph(xml, seg[0].index);
  }

  if (seg[0].type === 't') {
    if (seg.length === 1) return findTable(xml, seg[0].index);
    if (seg.length === 2 && seg[1].type === 'r') return findRow(xml, seg[0].index, seg[1].index);
    if (seg.length >= 3 && seg[1].type === 'r' && seg[2].type === 'c') {
      return findCell(xml, seg[0].index, seg[1].index, seg[2].index);
    }
  }

  throw new Error('Invalid path syntax: "' + path + '" — 올바른 형식: t0, t0.r1, t0.r1.c3, p0 등');
}

// ──────── 문서 요소 수집 ────────

/** SECTION 내 P(본문)와 TABLE을 문서 순서대로 수집 */
export function collectDocElements(section: string): DocElement[] {
  const elements: DocElement[] = [];

  // TABLE 수집 (중첩 안전)
  const marker = '{{TABLE}}';
  const { replaced: noTables, tables } = replaceTopLevelTables(section, marker);
  const tableMatches = tables.map(t => ({ start: t.index, length: t.match.length }));

  for (const t of tables) {
    elements.push({ type: 'table', index: t.index, content: t.match });
  }
  const pRe = /<P[\s\S]*?<\/P>/g;
  let pm;
  while ((pm = pRe.exec(noTables))) {
    if (pm[0].includes(marker)) continue;

    let origOffset = pm.index;
    let searchPos = 0;
    for (const t of tableMatches) {
      const markerPos = noTables.indexOf(marker, searchPos);
      if (markerPos < pm.index) {
        origOffset += (t.length - marker.length);
        searchPos = markerPos + marker.length;
      } else {
        break;
      }
    }

    elements.push({ type: 'p', index: origOffset, content: pm[0] });
  }

  elements.sort((a, b) => a.index - b.index);
  return elements;
}

/** 표를 CellInfo 3차원 배열로 파싱 */
export function parseTables(xml: string): CellInfo[][][] {
  const allTables: CellInfo[][][] = [];
  const topTables = matchTopLevelTables(xml);
  const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
  const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;

  for (const tm of topTables) {
    const rows: CellInfo[][] = [];
    let rm;
    rowRe.lastIndex = 0;
    while ((rm = rowRe.exec(tm.match))) {
      const cells: CellInfo[] = [];
      let cm;
      cellRe.lastIndex = 0;
      while ((cm = cellRe.exec(rm[0]))) {
        const cellTag = cm[0];
        const col = getAttr(cellTag, 'ColAddr');
        const row = getAttr(cellTag, 'RowAddr');
        const colSpan = getAttr(cellTag, 'ColSpan');
        const rowSpan = getAttr(cellTag, 'RowSpan');
        const text = extractText(cellTag);
        cells.push({
          col: col !== null ? col : 0,
          row: row !== null ? row : 0,
          colSpan: colSpan !== null ? colSpan : 1,
          rowSpan: rowSpan !== null ? rowSpan : 1,
          text,
        });
      }
      rows.push(cells);
    }
    allTables.push(rows);
  }
  return allTables;
}

/** CHAR 태그에서 텍스트 추출 */
export function extractText(cellXml: string): string {
  const texts: string[] = [];
  const charRe = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>/g;
  let m;
  while ((m = charRe.exec(cellXml))) {
    const t = unescapeXml(m[1]).trim();
    if (t) texts.push(t);
  }
  return texts.join(' ');
}
