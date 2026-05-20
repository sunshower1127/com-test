/**
 * GET 계열 — 문서 구조 읽기
 * 순수 함수: XML 입력 → 문자열 출력 (COM 호출 없음)
 */

import { extractSection, collectDocElements, parseTables, getAttr, parseRange, extractText, matchTopLevelTables } from './path-resolver';
import { compactParagraph, compactTable, cleanTable, cleanParagraph } from './xml-cleaner';
import { CellInfo, PageBoundary } from './types';

// ──────── outline — 경량 목차 ────────

/** 문서 목차 (크기만, 내용 없음). 드릴다운 + 범위 지정 가능 */
export function outline(xml: string, path?: string, pageBoundaries?: PageBoundary[]): string {
  if (!path) return outlineDocument(xml, pageBoundaries || []);

  const { start, end } = parseRange(path);
  const seg = start.segments;

  // 범위 우선 체크: t0.r0~t0.r3, t0.r1.c3~t0.r1.c5
  if (end) {
    return outlineRange(xml, start, end);
  }

  // 드릴다운 (단일): t0 → 표 내부, t0.r1 → 행 내부, t0.r1.c3 → 셀 내부
  if (seg[0].type === 't') {
    if (seg.length === 1) return outlineTable(xml, seg[0].index);
    if (seg.length === 2 && seg[1].type === 'r') return outlineRow(xml, seg[0].index, seg[1].index);
    if (seg.length === 3 && seg[1].type === 'r' && seg[2].type === 'c') {
      return outlineCell(xml, seg[0].index, seg[1].index, seg[2].index);
    }
    if (seg.length === 4 && seg[1].type === 'r' && seg[2].type === 'c' && seg[3].type === 'p') {
      return outlineParagraph(xml, seg[0].index, seg[1].index, seg[2].index, seg[3].index);
    }
  }

  // 본문 p: p0
  if (seg.length === 1 && seg[0].type === 'p') {
    return outlineBodyParagraph(xml, seg[0].index);
  }

  return 'Unknown outline path: ' + path;
}

function outlineDocument(xml: string, pageBoundaries: PageBoundary[]): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);

  // 각 요소의 정보 수집
  interface OutlineItem {
    id: string;
    type: 'p' | 'table';
    line: string;
    // 본문 P: para 번호 (List=0 기준)
    paraIdx?: number;
    // TABLE: 첫 셀과 마지막 셀의 List ID 범위 (추정)
    startList?: number;
    endList?: number;
  }
  const items: OutlineItem[] = [];
  let tIdx = 0;
  // bodyParaCount = COM Para 인덱스 (P + TABLE 모두 카운트)
  let bodyParaCount = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      const textLen = countTextLength(el.content);
      // 빈 P는 outline에서 생략 (bodyParaCount는 항상 증가)
      if (textLen > 0) {
        items.push({
          id: 'p' + bodyParaCount,
          type: 'p',
          line: '<P id="p' + bodyParaCount + '" len=' + textLen + '/>',
          paraIdx: bodyParaCount,
        });
      }
      bodyParaCount++;
    } else {
      const compactLen = compactTable(el.content, tIdx, {}).length;
      const rowCount = el.content.match(/RowCount="(\d+)"/);
      const colCount = el.content.match(/ColCount="(\d+)"/);
      const rows = rowCount ? rowCount[1] : '?';
      const cols = colCount ? colCount[1] : '?';
      items.push({
        id: 't' + tIdx,
        type: 'table',
        line: '<TABLE id="t' + tIdx + '" rows=' + rows + ' cols=' + cols + ' len=' + compactLen + '/>',
        paraIdx: bodyParaCount,
      });
      tIdx++;
      // 표는 본문에서 P 하나를 차지 (표 컨트롤이 P 안에 있음)
      bodyParaCount++;
    }
  }

  // 페이지 정보가 없으면 그냥 출력
  if (pageBoundaries.length === 0) {
    return items.map(item => item.line).join('\n');
  }

  // 페이지 경계 기반 배정
  // pageBoundaries[i].startPara ~ endPara: 해당 페이지의 본문(List=0) 문단 범위
  const out: string[] = [];
  let currentPage = -1;
  const shownIds = new Set<string>();

  for (const item of items) {
    // 이 요소가 어느 페이지에 속하는지 판별
    const para = item.paraIdx !== undefined ? item.paraIdx : -1;
    let itemPage = 0;

    for (let pi = 0; pi < pageBoundaries.length; pi++) {
      const pb = pageBoundaries[pi];
      // 본문(List=0) 기준: startPara <= para <= endPara
      if (pb.startList === 0 && para >= pb.startPara && para <= pb.endPara) {
        itemPage = pi;
        break;
      }
      // 표 안(List>0)이거나 경계를 넘은 경우: endPara 이후면 다음 페이지
      if (pb.startList === 0 && para > pb.endPara) {
        itemPage = pi + 1;
      }
    }
    itemPage = Math.min(itemPage, pageBoundaries.length - 1);

    // 새 페이지 시작
    if (itemPage > currentPage) {
      // 이전 페이지에서 시작된 표가 이 페이지에도 걸치는지 체크
      // (직전 요소가 TABLE이고 이 페이지에도 영향이 있으면 continues)
      currentPage = itemPage;
      out.push('=== Page ' + (currentPage + 1) + ' ===');
    }

    if (shownIds.has(item.id)) {
      out.push('<' + (item.type === 'table' ? 'TABLE' : 'P') + ' id="' + item.id + '" continues/>');
    } else {
      out.push(item.line);
      shownIds.add(item.id);
    }
  }

  return out.join('\n');
}

function outlineTable(xml: string, tableIdx: number): string {
  const tables = parseTables(xml);
  const table = tables[tableIdx];
  if (!table) return 'Table not found: t' + tableIdx;

  // 원본 XML에서 셀별 상세 정보
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  const tableXml = tableIdx < topTables.length ? topTables[tableIdx].match : '';

  const rowCount = tableXml.match(/RowCount="(\d+)"/);
  const colCount = tableXml.match(/ColCount="(\d+)"/);
  const rows = rowCount ? rowCount[1] : '?';
  const cols = colCount ? colCount[1] : '?';
  const out: string[] = [];
  out.push('<TABLE id="t' + tableIdx + '" rows=' + rows + ' cols=' + cols + '>');

  const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
  let rm;
  let ri = 0;
  while ((rm = rowRe.exec(tableXml))) {
    out.push('  <ROW>');
    const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
    let cm;
    while ((cm = cellRe.exec(rm[0]))) {
      const colAddr = getAttr(cm[0], 'ColAddr');
      const colSpan = getAttr(cm[0], 'ColSpan') || 1;
      const rowSpan = getAttr(cm[0], 'RowSpan') || 1;
      const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';
      const path = 't' + tableIdx + '.r' + ri + '.c' + colAddr;

      const textLen = countTextLength(cm[2]);
      const imgCount = (cm[2].match(/<PICTURE/g) || []).length;
      const imgAttr = imgCount > 0 ? ' img=' + imgCount : '';

      out.push('    <CELL id="' + path + '"' + spanAttr + ' len=' + textLen + imgAttr + '/>');
    }
    out.push('  </ROW>');
    ri++;
  }

  out.push('</TABLE>');
  return out.join('\n');
}

function outlineRow(xml: string, tableIdx: number, rowIdx: number): string {
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  for (let ti = 0; ti < topTables.length; ti++) {
    if (ti === tableIdx) {
      const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(topTables[ti].match))) {
        if (ri === rowIdx) {
          const out: string[] = ['<ROW>'];
          const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
          let cm;
          while ((cm = cellRe.exec(rm[0]))) {
            const colAddr = getAttr(cm[0], 'ColAddr');
            const colSpan = getAttr(cm[0], 'ColSpan') || 1;
            const rowSpan = getAttr(cm[0], 'RowSpan') || 1;
            const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';
            const path = 't' + tableIdx + '.r' + rowIdx + '.c' + colAddr;
            const textLen = countTextLength(cm[2]);
            const imgCount = (cm[2].match(/<PICTURE/g) || []).length;
            const imgAttr = imgCount > 0 ? ' img=' + imgCount : '';
            out.push('  <CELL id="' + path + '"' + spanAttr + ' len=' + textLen + imgAttr + '/>');
          }
          out.push('</ROW>');
          return out.join('\n');
        }
        ri++;
      }
    }
  }
  return 'Row not found: t' + tableIdx + '.r' + rowIdx;
}

function outlineCell(xml: string, tableIdx: number, rowIdx: number, colIdx: number): string {
  // 셀 XML 찾기
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  if (tableIdx < topTables.length) {
    const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
    let rm;
    let ri = 0;
    while ((rm = rowRe.exec(topTables[tableIdx].match))) {
      if (ri === rowIdx) {
        const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
        let cm;
        while ((cm = cellRe.exec(rm[0]))) {
          const ca = getAttr(cm[0], 'ColAddr');
          if (ca === colIdx) {
            const out: string[] = [];
            const path = 't' + tableIdx + '.r' + rowIdx + '.c' + colIdx;
            out.push('<CELL id="' + path + '">');

            // P 태그별 길이 (빈 P는 생략)
            const pRe = /<P[\s\S]*?<\/P>/g;
            let pm;
            let pi = 0;
            while ((pm = pRe.exec(cm[0]))) {
              const pLen = countTextLength(pm[0]);
              if (pLen > 0) out.push('  <P id="' + path + '.p' + pi + '" len=' + pLen + '/>');
              pi++;
            }

            out.push('</CELL>');
            return out.join('\n');
          }
        }
      }
      ri++;
    }
  }
  return 'Cell not found: 지정한 경로의 셀을 찾을 수 없습니다';
}

function outlineParagraph(xml: string, tableIdx: number, rowIdx: number, colIdx: number, paraIdx: number): string {
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  if (tableIdx < topTables.length) {
    const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
    let rm;
    let ri = 0;
    while ((rm = rowRe.exec(topTables[tableIdx].match))) {
      if (ri === rowIdx) {
        const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
        let cm;
        while ((cm = cellRe.exec(rm[0]))) {
          const ca = getAttr(cm[0], 'ColAddr');
          if (ca === colIdx) {
            const pRe = /<P[\s\S]*?<\/P>/g;
            let pm;
            let pi = 0;
            while ((pm = pRe.exec(cm[0]))) {
              if (pi === paraIdx) {
                const pLen = countTextLength(pm[0]);
                const id = 't' + tableIdx + '.r' + rowIdx + '.c' + colIdx + '.p' + paraIdx;
                return '<P id="' + id + '" len=' + pLen + '/>';
              }
              pi++;
            }
            return 'Paragraph not found: t' + tableIdx + '.r' + rowIdx + '.c' + colIdx + '.p' + paraIdx + ' (문단 ' + paraIdx + ' 없음, 최대 ' + pi + '개)';
          }
        }
        return 'Cell not found: 지정한 경로의 셀을 찾을 수 없습니다';
      }
      ri++;
    }
  }
  return 'Table not found: t' + tableIdx + ' (표 ' + tableIdx + ' 없음, 최대 ' + topTables.length + '개)';
}

function outlineBodyParagraph(xml: string, paraIdx: number): string {
  const { section } = extractSection(xml);
  const elems = collectDocElements(section);
  let bodyIdx = 0;
  for (const el of elems) {
    if (bodyIdx === paraIdx) {
      if (el.type !== 'p') return 'p' + paraIdx + ' 위치는 TABLE입니다';
      const pLen = countTextLength(el.content);
      return '<P id="p' + paraIdx + '" len=' + pLen + '/>';
    }
    bodyIdx++;
  }
  return 'Paragraph not found: p' + paraIdx + ' (문단 ' + paraIdx + ' 없음, 최대 ' + bodyIdx + '개)';
}

function outlineRange(xml: string, start: ReturnType<typeof parseRange>['start'], end: NonNullable<ReturnType<typeof parseRange>['end']>): string {
  const startSeg = start.segments;
  const endSeg = end.segments;

  // t0.r1.c3~t0.r1.c5: 같은 행 내 셀 범위
  if (startSeg.length >= 3 && endSeg.length >= 3 &&
      startSeg[0].type === 't' && startSeg[1].type === 'r' && startSeg[2].type === 'c') {
    const tableIdx = startSeg[0].index;
    const rowIdx = startSeg[1].index;
    const fromCol = startSeg[2].index;
    const toCol = endSeg[2].index;

    const { section } = extractSection(xml);
    const topTables = matchTopLevelTables(section);
    if (tableIdx < topTables.length) {
      const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
      let rm; let ri = 0;
      while ((rm = rowRe.exec(topTables[tableIdx].match))) {
        if (ri === rowIdx) {
          const out: string[] = [];
          const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
          let cm;
          while ((cm = cellRe.exec(rm[0]))) {
            const colAddr = getAttr(cm[0], 'ColAddr') || 0;
            if (colAddr >= fromCol && colAddr <= toCol) {
              const colSpan = getAttr(cm[0], 'ColSpan') || 1;
              const rowSpan = getAttr(cm[0], 'RowSpan') || 1;
              const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';
              const path = 't' + tableIdx + '.r' + rowIdx + '.c' + colAddr;
              const textLen = countTextLength(cm[2]);
              const imgCount = (cm[2].match(/<PICTURE/g) || []).length;
              const imgAttr = imgCount > 0 ? ' img=' + imgCount : '';
              out.push('<CELL id="' + path + '"' + spanAttr + ' len=' + textLen + imgAttr + '/>');
            }
          }
          return out.join('\n');
        }
        ri++;
      }
    }
    return 'Range not found';
  }

  // t0.r0~t0.r3: 행 범위
  if (startSeg.length >= 2 && endSeg.length >= 2 &&
      startSeg[0].type === 't' && startSeg[1].type === 'r') {
    const tableIdx = startSeg[0].index;
    const fromRow = startSeg[1].index;
    const toRow = endSeg[1].index;

    const { section } = extractSection(xml);
    const topTables = matchTopLevelTables(section);
    if (tableIdx < topTables.length) {
      const out: string[] = [];
      const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
      let rm; let ri = 0;
      while ((rm = rowRe.exec(topTables[tableIdx].match))) {
          if (ri >= fromRow && ri <= toRow) {
            out.push('<ROW>');
            const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
            let cm;
            while ((cm = cellRe.exec(rm[0]))) {
              const colAddr = getAttr(cm[0], 'ColAddr');
              const colSpan = getAttr(cm[0], 'ColSpan') || 1;
              const rowSpan = getAttr(cm[0], 'RowSpan') || 1;
              const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';
              const path = 't' + tableIdx + '.r' + ri + '.c' + colAddr;
              const textLen = countTextLength(cm[2]);
              const imgCount = (cm[2].match(/<PICTURE/g) || []).length;
              const imgAttr = imgCount > 0 ? ' img=' + imgCount : '';
              out.push('  <CELL id="' + path + '"' + spanAttr + ' len=' + textLen + imgAttr + '/>');
            }
            out.push('</ROW>');
          }
          ri++;
        }
      return out.join('\n');
    }
    return 'Range not found';
  }

  // p0~p5: 본문 범위 (bodyIdx = COM Para)
  if (startSeg[0].type === 'p' && endSeg[0].type === 'p') {
    const from = startSeg[0].index;
    const to = endSeg[0].index;
    const { section } = extractSection(xml);
    const elements = collectDocElements(section);
    const out: string[] = [];
    let bodyIdx = 0;
    for (const el of elements) {
      if (bodyIdx >= from && bodyIdx <= to && el.type === 'p') {
        const textLen = countTextLength(el.content);
        out.push('<P id="p' + bodyIdx + '" len=' + textLen + '/>');
      }
      bodyIdx++;
    }
    return out.join('\n');
  }

  return 'Range not supported';
}

/** CHAR 태그에서 순수 텍스트 길이 세기 */
function countTextLength(xml: string): number {
  let len = 0;
  const charRe = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>/g;
  let m;
  while ((m = charRe.exec(xml))) {
    len += m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').length;
  }
  return len;
}

// ──────── get — 간결 구조맵 ────────

/** 간결 구조맵 (범위 지정 가능) */
export function get(xml: string, listIdMap: Record<string, number>, path?: string): string {
  if (!path) return structure(xml, listIdMap);

  const hints = parseCharStyleHints(xml);
  const { start, end } = parseRange(path);
  const seg = start.segments;

  // 단일 요소
  if (!end) {
    if (seg[0].type === 'p') {
      return getBodyParagraph(xml, seg[0].index, hints);
    }
    if (seg[0].type === 't') {
      if (seg.length === 1) return getTable(xml, seg[0].index, listIdMap, hints);
      if (seg.length === 2 && seg[1].type === 'r') return getRow(xml, seg[0].index, seg[1].index, listIdMap, hints);
      if (seg.length >= 3) {
        const paraIdx = seg.length >= 4 && seg[3].type === 'p' ? seg[3].index : undefined;
        return getCell(xml, seg[0].index, seg[1].index, seg[2].index, listIdMap, paraIdx, hints);
      }
    }
  }

  // 범위
  if (end) {
    return getRange(xml, start, end, listIdMap, hints);
  }

  return 'Unknown path: ' + path;
}

function getBodyParagraph(xml: string, paraIdx: number, hints?: Record<number, CharStyleHint>): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  let bodyIdx = 0;
  for (const el of elements) {
    if (bodyIdx === paraIdx) {
      if (el.type !== 'p') return 'p' + paraIdx + ' 위치는 TABLE입니다';
      if (hints) {
        const text = extractStyledText(el.content, hints);
        return '<P id="p' + bodyIdx + '">' + text + '</P>';
      }
      return compactParagraph(el.content, bodyIdx);
    }
    bodyIdx++;
  }
  return 'Paragraph not found: p' + paraIdx;
}

function getTable(xml: string, tableIdx: number, listIdMap: Record<string, number>, hints?: Record<number, CharStyleHint>): string {
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  if (tableIdx < topTables.length) return compactTable(topTables[tableIdx].match, tableIdx, listIdMap);
  return 'Table not found: t' + tableIdx;
}

function getRow(xml: string, tableIdx: number, rowIdx: number, listIdMap: Record<string, number>, hints?: Record<number, CharStyleHint>): string {
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  if (tableIdx < topTables.length) {
    const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
    let rm;
    let ri = 0;
    while ((rm = rowRe.exec(topTables[tableIdx].match))) {
      if (ri === rowIdx) {
        const out: string[] = ['<ROW>'];
        const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
        let cm;
        while ((cm = cellRe.exec(rm[0]))) {
          const colAddr = getAttr(cm[0], 'ColAddr');
          const colSpan = getAttr(cm[0], 'ColSpan') || 1;
          const rowSpan = getAttr(cm[0], 'RowSpan') || 1;
          const path = 't' + tableIdx + '.r' + ri + '.c' + colAddr;
          const listId = listIdMap[path] !== undefined ? ' L=' + listIdMap[path] : '';
          const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';

          const content = formatCellContent(cm[0], path, '    ', hints);

          out.push('  <CELL id="' + path + '"' + spanAttr + listId + '>' + content + '</CELL>');
        }
        out.push('</ROW>');
        return out.join('\n');
      }
      ri++;
    }
  }
  return 'Row not found: t' + tableIdx + '.r' + rowIdx;
}

function getCell(xml: string, tableIdx: number, rowIdx: number, colIdx: number, listIdMap: Record<string, number>, paraIdx?: number, hints?: Record<number, CharStyleHint>): string {
  const { section } = extractSection(xml);
  const topTables = matchTopLevelTables(section);
  if (tableIdx < topTables.length) {
    const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
    let rm;
    let ri = 0;
    while ((rm = rowRe.exec(topTables[tableIdx].match))) {
      if (ri === rowIdx) {
        const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
        let cm;
        while ((cm = cellRe.exec(rm[0]))) {
          const ca = getAttr(cm[0], 'ColAddr');
          if (ca === colIdx) {
            const path = 't' + tableIdx + '.r' + rowIdx + '.c' + colIdx;
            const listId = listIdMap[path] !== undefined ? ' L=' + listIdMap[path] : '';

            // 특정 P만 요청된 경우
            if (paraIdx !== undefined) {
              const pRe = /<P[\s\S]*?<\/P>/g;
              let pMatch;
              let pi = 0;
              while ((pMatch = pRe.exec(cm[0]))) {
                if (pi === paraIdx) {
                  const text = hints ? extractStyledText(pMatch[0], hints) : extractInlineText(pMatch[0]);
                  return '<P id="' + path + '.p' + paraIdx + '">' + text + '</P>';
                }
                pi++;
              }
              return 'P not found: ' + path + '.p' + paraIdx;
            }

            const content = formatCellContent(cm[0], path, '  ', hints);
            return '<CELL id="' + path + '"' + listId + '>' + content + '</CELL>';
          }
        }
      }
      ri++;
    }
  }
  return 'Cell not found: 지정한 경로의 셀을 찾을 수 없습니다';
}

function getRange(xml: string, start: ReturnType<typeof parseRange>['start'], end: NonNullable<ReturnType<typeof parseRange>['end']>, listIdMap: Record<string, number>, hints?: Record<number, CharStyleHint>): string {
  const startSeg = start.segments;
  const endSeg = end.segments;

  // p0~p2: 본문 문단 범위
  if (startSeg[0].type === 'p' && endSeg[0].type === 'p') {
    const from = startSeg[0].index;
    const to = endSeg[0].index;
    const { section } = extractSection(xml);
    const elements = collectDocElements(section);
    const out: string[] = [];
    let bodyIdx = 0;
    for (const el of elements) {
      if (bodyIdx >= from && bodyIdx <= to && el.type === 'p') {
        out.push(compactParagraph(el.content, bodyIdx));
      }
      bodyIdx++;
    }
    return out.join('\n');
  }

  // t0.r0~t0.r3: 표 행 범위
  if (startSeg.length >= 2 && startSeg[1].type === 'r' && endSeg.length >= 2 && endSeg[1].type === 'r') {
    const tableIdx = startSeg[0].index;
    const fromRow = startSeg[1].index;
    const toRow = endSeg[1].index;
    const out: string[] = [];

    const { section } = extractSection(xml);
    const topTables = matchTopLevelTables(section);
    if (tableIdx < topTables.length) {
      const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(topTables[tableIdx].match))) {
        if (ri >= fromRow && ri <= toRow) {
          const rowOut: string[] = ['  <ROW>'];
          const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
          let cm;
          while ((cm = cellRe.exec(rm[0]))) {
            const colAddr = getAttr(cm[0], 'ColAddr');
            const colSpan = getAttr(cm[0], 'ColSpan') || 1;
            const rowSpan = getAttr(cm[0], 'RowSpan') || 1;
            const path = 't' + tableIdx + '.r' + ri + '.c' + colAddr;
            const listId = listIdMap[path] !== undefined ? ' L=' + listIdMap[path] : '';
            const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';
            const content = formatCellContent(cm[0], path, '      ', hints);
            rowOut.push('    <CELL id="' + path + '"' + spanAttr + listId + '>' + content + '</CELL>');
          }
          rowOut.push('  </ROW>');
          out.push(rowOut.join('\n'));
        }
        ri++;
      }
    }
    return out.join('\n');
  }

  return 'Range not supported: ' + JSON.stringify({ start: startSeg, end: endSeg });
}

/** XML에서 인라인 텍스트 추출 (CHAR + LINEBREAK) */
/** 셀 내용 포맷팅 — P 1개면 인라인, 여러 개면 id 붙이고 빈 P 제거 */
function formatCellContent(cellXml: string, path: string, indent: string = '  ', hints?: Record<number, CharStyleHint>): string {
  const hasImage = cellXml.includes('<PICTURE');
  if (hasImage) return '[IMAGE]';

  const extract = hints ? (x: string) => extractStyledText(x, hints) : extractInlineText;

  const pMatches: string[] = [];
  const pRe = /<P[\s\S]*?<\/P>/g;
  let pm;
  while ((pm = pRe.exec(cellXml))) pMatches.push(pm[0]);

  if (pMatches.length <= 1) {
    return extract(cellXml);
  }

  // P 여러 개: id 붙이고 빈 P 제거
  const pLines: string[] = [];
  pMatches.forEach((pxml, pi) => {
    const text = extract(pxml);
    if (text) {
      pLines.push('\n' + indent + '<P id="' + path + '.p' + pi + '">' + text + '</P>');
    }
  });
  return pLines.length > 0 ? pLines.join('') + '\n' + indent.substring(2) : '';
}

function extractInlineText(xml: string): string {
  const parts: string[] = [];
  const re = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<LINEBREAK\s*\/?>/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('<LINEBREAK')) {
      parts.push('\\n');
    } else if (m[1]) {
      const t = m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/\n/g, '\\n').replace(/\r/g, '');
      parts.push(t);
    }
  }
  return parts.join('');
}

// ──────── CharShape 특성맵 (스타일 마커용) ────────

interface CharStyleHint {
  bold: boolean;
  italic: boolean;
  color: string | null; // null = 검정(기본), '#FF0000' 등
}

/** HEAD에서 CharShape별 bold/italic/color 특성 파싱 */
function parseCharStyleHints(xml: string): Record<number, CharStyleHint> {
  const hints: Record<number, CharStyleHint> = {};
  const headMatch = xml.match(/<HEAD[^>]*>([\s\S]*?)<\/HEAD>/);
  if (!headMatch) return hints;
  const head = headMatch[1];
  const re = /<CHARSHAPE\b([^>]*)(?:\/>|>([\s\S]*?)<\/CHARSHAPE>)/g;
  let m;
  let idx = 0;
  while ((m = re.exec(head))) {
    const attrs = m[1];
    const children = m[2] || '';
    const bold = children.includes('<BOLD');
    const italic = children.includes('<ITALIC');
    const colorMatch = /TextColor="(\d+)"/.exec(attrs);
    const colorVal = colorMatch ? parseInt(colorMatch[1]) : 0;
    const color = colorVal !== 0 ? colorToHex(String(colorVal)) : null;
    if (bold || italic || color) {
      hints[idx] = { bold, italic, color };
    }
    idx++;
  }
  return hints;
}

/** TEXT의 CharShape에 따라 스타일 마커를 감싸서 텍스트 추출 */
function extractStyledText(xml: string, hints: Record<number, CharStyleHint>): string {
  const parts: string[] = [];
  // TEXT 단위로 순회
  const textRe = /<TEXT\b[^>]*CharShape="(\d+)"[^>]*(?:\/>|>([\s\S]*?)<\/TEXT>)/g;
  let tm;
  while ((tm = textRe.exec(xml))) {
    const csId = parseInt(tm[1]);
    const content = tm[2] || '';
    // TEXT 안의 CHAR/LINEBREAK 추출
    const text = extractCharsFromText(content);
    if (!text) continue;
    const hint = hints[csId];
    if (hint) {
      // 태그명 조합: b, i, bi
      let tag = '';
      if (hint.bold) tag += 'b';
      if (hint.italic) tag += 'i';
      if (!tag && hint.color) tag = 'c';
      const colorAttr = hint.color ? ' ' + hint.color : '';
      parts.push('<' + (tag || 'c') + colorAttr + '>' + text + '</' + (tag || 'c') + '>');
    } else {
      parts.push(text);
    }
  }
  return parts.join('');
}

/** TEXT 내부의 CHAR/LINEBREAK만 추출 (스타일 없이) */
function extractCharsFromText(content: string): string {
  const parts: string[] = [];
  const re = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<CHAR\b[^>]*\/>|<LINEBREAK\s*\/?>/g;
  let m;
  while ((m = re.exec(content))) {
    if (m[0].startsWith('<LINEBREAK')) {
      parts.push('\\n');
    } else if (m[1]) {
      const t = m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/\n/g, '\\n').replace(/\r/g, '');
      parts.push(t);
    }
  }
  return parts.join('');
}

// ──────── styles — HEAD 섹션 스타일 조회 ────────

/** CharShape/ParaShape 조회 */
// API名 → HWPML2Xタグ名 マッピング
const STYLE_TAG_MAP: Record<string, string> = {
  CharShape: 'CHARSHAPE',
  ParaShape: 'PARASHAPE',
  BorderFill: 'BORDERFILL',
};

export function styles(xml: string, type?: string, id?: number): string {
  const headMatch = xml.match(/<HEAD[^>]*>([\s\S]*?)<\/HEAD>/);
  if (!headMatch) return 'HEAD section not found';

  const head = headMatch[1];
  const fontMap = parseFontMap(head);
  const out: string[] = [];

  if (!type || type === 'CharShape') {
    out.push(formatStyleList(head, 'CHARSHAPE', 'CharShape', fontMap, id));
  }
  if (!type || type === 'ParaShape') {
    out.push(formatStyleList(head, 'PARASHAPE', 'ParaShape', fontMap, id));
  }
  if (type && type !== 'CharShape' && type !== 'ParaShape') {
    const xmlTag = STYLE_TAG_MAP[type] || type.toUpperCase();
    out.push(formatStyleList(head, xmlTag, type, fontMap, id));
  }

  return out.join('\n\n');
}

/** FACENAMELIST에서 Hangul 폰트 ID→이름 매핑 */
function parseFontMap(head: string): Record<number, string> {
  const map: Record<number, string> = {};
  // Hangul 폰트페이스 블록 찾기
  const faceRe = /<FONTFACE[^>]*Lang="Hangul"[^>]*>([\s\S]*?)<\/FONTFACE>/;
  const fm = faceRe.exec(head);
  if (!fm) return map;
  const fontRe = /<FONT[^>]*Id="(\d+)"[^>]*Name="([^"]*)"[^>]*\/?>/g;
  let m;
  while ((m = fontRe.exec(fm[1]))) {
    map[parseInt(m[1])] = m[2];
  }
  return map;
}

function formatStyleList(head: string, xmlTag: string, displayName: string, fontMap: Record<number, string>, specificId?: number): string {
  // self-closing과 자식 요소 포함 모두 매칭
  const re = new RegExp('<' + xmlTag + '\\b([^>]*)(?:/>|>([\\s\\S]*?)</' + xmlTag + '>)', 'g');
  const out: string[] = [];
  let m;
  let idx = 0;

  if (specificId !== undefined) {
    // 특정 ID 상세
    while ((m = re.exec(head))) {
      if (idx === specificId) {
        out.push(displayName + '[' + idx + ']:');
        const attrs = parseAllAttrs(m[1]);
        for (const [k, v] of Object.entries(attrs)) {
          out.push('  ' + k + ': ' + v);
        }
        // 자식 태그도 표시
        const children = m[2] || '';
        const childTags = parseChildTags(children);
        for (const [tag, childAttrs] of childTags) {
          if (childAttrs) {
            out.push('  ' + tag + ': ' + childAttrs);
          } else {
            out.push('  ' + tag + ': ✓');
          }
        }
        return out.join('\n');
      }
      idx++;
    }
    return displayName + '[' + specificId + ']: not found';
  }

  // 전체 목록 (요약)
  out.push(displayName + ':');
  while ((m = re.exec(head))) {
    const attrs = parseAllAttrs(m[1]);
    const children = m[2] || '';
    const summary = summarizeStyle(displayName, attrs, children, fontMap);
    out.push('  ' + idx + ': ' + summary);
    idx++;
  }

  if (idx === 0) out.push('  (none)');
  return out.join('\n');
}

/** 자식 태그 파싱 — [태그명, 속성요약] 배열 반환 */
function parseChildTags(children: string): Array<[string, string]> {
  const result: Array<[string, string]> = [];
  const re = /<([A-Z]+)\b([^>]*)\/?>/g;
  let m;
  while ((m = re.exec(children))) {
    const tag = m[1];
    // FONTID, RATIO 등 기본 구성 요소는 건너뜀
    if (['FONTID', 'RATIO', 'CHARSPACING', 'RELSIZE', 'CHAROFFSET'].includes(tag)) continue;
    const attrs = m[2].trim();
    result.push([tag, attrs ? attrs.replace(/"/g, '').replace(/\s+/g, ' ') : '']);
  }
  return result;
}

function summarizeStyle(tagName: string, attrs: Record<string, string>, children?: string, fontMap?: Record<number, string>): string {
  if (tagName === 'CharShape') {
    const parts: string[] = [];
    // 폰트명 (FONTID Hangul 참조)
    if (fontMap && children) {
      const fidMatch = /FONTID[^>]*Hangul="(\d+)"/.exec(children);
      if (fidMatch) {
        const fontName = fontMap[parseInt(fidMatch[1])];
        if (fontName) parts.push(fontName);
      }
    }
    if (attrs['Height']) parts.push(Math.round(parseInt(attrs['Height']) / 100) + 'pt');
    // 자식 태그에서 Bold/Italic/Underline/StrikeOut 감지
    const body = children || '';
    if (body.includes('<BOLD')) parts.push('볼드');
    if (body.includes('<ITALIC')) parts.push('이탤릭');
    if (body.includes('<UNDERLINE')) parts.push('밑줄');
    if (body.includes('<STRIKEOUT')) parts.push('취소선');
    if (body.includes('<SUPERSCRIPT')) parts.push('위첨자');
    if (body.includes('<SUBSCRIPT')) parts.push('아래첨자');
    if (attrs['TextColor'] && attrs['TextColor'] !== '0') parts.push('색:' + colorToHex(attrs['TextColor']));
    return parts.join(' ') || '기본';
  }

  if (tagName === 'ParaShape') {
    const parts: string[] = [];
    if (attrs['Align']) {
      const alignMap: Record<string, string> = { '0': '양쪽', '1': '왼쪽', '2': '오른쪽', '3': '가운데' };
      parts.push(alignMap[attrs['Align']] || attrs['Align']);
    }
    if (attrs['LineSpacing']) parts.push('줄간격' + attrs['LineSpacing'] + '%');
    return parts.join(' ') || '기본';
  }

  return Object.entries(attrs).slice(0, 3).map(([k, v]) => k + '=' + v).join(' ');
}

/** Windows COLORREF (10진수) → #RRGGBB */
function colorToHex(val: string): string {
  const n = parseInt(val);
  if (isNaN(n) || n === 0) return '#000000';
  const r = n & 0xFF;
  const g = (n >> 8) & 0xFF;
  const b = (n >> 16) & 0xFF;
  return '#' + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1).toUpperCase();
}

function parseAllAttrs(attrStr: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /(\w+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(attrStr))) {
    attrs[m[1]] = m[2];
  }
  return attrs;
}

// ──────── 기존 호환 ────────

/** 기존 structure (= get 전체) */
export function structure(xml: string, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  const out: string[] = [];
  let bodyIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      out.push(compactParagraph(el.content, bodyIdx));
    } else {
      out.push(compactTable(el.content, tIdx, listIdMap));
      tIdx++;
    }
    bodyIdx++;
  }

  return out.join('\n');
}

/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 */
export function structureDetail(xml: string, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  const out: string[] = [];
  let bodyIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      out.push(cleanParagraph(el.content, 'p' + bodyIdx, 0));
    } else {
      out.push(cleanTable(el.content, tIdx, 0, listIdMap));
      tIdx++;
    }
    bodyIdx++;
  }

  return out.join('\n');
}
