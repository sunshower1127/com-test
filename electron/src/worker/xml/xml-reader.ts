/**
 * GET 계열 — 문서 구조 읽기
 * 순수 함수: XML 입력 → 문자열 출력 (COM 호출 없음)
 */

import { extractSection, collectDocElements, parseTables, getAttr, parseRange, extractText } from './path-resolver';
import { compactParagraph, compactTable, cleanTable, cleanParagraph } from './xml-cleaner';
import { CellInfo } from './types';

// ──────── outline — 경량 목차 ────────

/** 문서 목차 (크기만, 내용 없음). 드릴다운 + 범위 지정 가능 */
export function outline(xml: string, path?: string): string {
  if (!path) return outlineDocument(xml);

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
    if (seg.length >= 3 && seg[1].type === 'r' && seg[2].type === 'c') {
      return outlineCell(xml, seg[0].index, seg[1].index, seg[2].index);
    }
  }

  return 'Unknown outline path: ' + path;
}

function outlineDocument(xml: string): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  const out: string[] = [];
  let pIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      const textLen = countTextLength(el.content);
      out.push('<P id="p' + pIdx + '" len=' + textLen + '/>');
      pIdx++;
    } else {
      const compactLen = compactTable(el.content, tIdx, {}).length;
      const rowCount = el.content.match(/RowCount="(\d+)"/);
      const colCount = el.content.match(/ColCount="(\d+)"/);
      const rows = rowCount ? rowCount[1] : '?';
      const cols = colCount ? colCount[1] : '?';
      out.push('<TABLE id="t' + tIdx + '" rows=' + rows + ' cols=' + cols + ' len=' + compactLen + '/>');
      tIdx++;
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
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  let tableXml = '';
  while ((tm = tableRe.exec(section))) {
    if (ti === tableIdx) { tableXml = tm[0]; break; }
    ti++;
  }

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
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(section))) {
    if (ti === tableIdx) {
      const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(tm[0]))) {
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
    ti++;
  }
  return 'Row not found: t' + tableIdx + '.r' + rowIdx;
}

function outlineCell(xml: string, tableIdx: number, rowIdx: number, colIdx: number): string {
  // 셀 XML 찾기
  const { section } = extractSection(xml);
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(section))) {
    if (ti === tableIdx) {
      const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(tm[0]))) {
        if (ri === rowIdx) {
          const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
          let cm;
          while ((cm = cellRe.exec(rm[0]))) {
            const ca = getAttr(cm[0], 'ColAddr');
            if (ca === colIdx) {
              const out: string[] = [];
              const path = 't' + tableIdx + '.r' + rowIdx + '.c' + colIdx;
              out.push('<CELL id="' + path + '">');

              // P 태그별 길이
              const pRe = /<P[\s\S]*?<\/P>/g;
              let pm;
              while ((pm = pRe.exec(cm[0]))) {
                const pLen = countTextLength(pm[0]);
                out.push('  <P len=' + pLen + '/>');
              }

              out.push('</CELL>');
              return out.join('\n');
            }
          }
        }
        ri++;
      }
    }
    ti++;
  }
  return 'Cell not found';
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
    const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
    let tm; let ti = 0;
    while ((tm = tableRe.exec(section))) {
      if (ti === tableIdx) {
        const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
        let rm; let ri = 0;
        while ((rm = rowRe.exec(tm[0]))) {
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
      ti++;
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
    const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
    let tm; let ti = 0;
    while ((tm = tableRe.exec(section))) {
      if (ti === tableIdx) {
        const out: string[] = [];
        const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
        let rm; let ri = 0;
        while ((rm = rowRe.exec(tm[0]))) {
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
      ti++;
    }
    return 'Range not found';
  }

  // p0~p2: 본문 문단 범위
  if (startSeg[0].type === 'p' && endSeg[0].type === 'p') {
    const from = startSeg[0].index;
    const to = endSeg[0].index;
    const { section } = extractSection(xml);
    const elements = collectDocElements(section);
    const out: string[] = [];
    let pi = 0;
    for (const el of elements) {
      if (el.type === 'p') {
        if (pi >= from && pi <= to) {
          const textLen = countTextLength(el.content);
          out.push('<P id="p' + pi + '" len=' + textLen + '/>');
        }
        pi++;
      }
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

  const { start, end } = parseRange(path);
  const seg = start.segments;

  // 단일 요소
  if (!end) {
    if (seg[0].type === 'p') {
      return getBodyParagraph(xml, seg[0].index);
    }
    if (seg[0].type === 't') {
      if (seg.length === 1) return getTable(xml, seg[0].index, listIdMap);
      if (seg.length === 2 && seg[1].type === 'r') return getRow(xml, seg[0].index, seg[1].index, listIdMap);
      if (seg.length >= 3) return getCell(xml, seg[0].index, seg[1].index, seg[2].index, listIdMap);
    }
  }

  // 범위
  if (end) {
    return getRange(xml, start, end, listIdMap);
  }

  return 'Unknown path: ' + path;
}

function getBodyParagraph(xml: string, paraIdx: number): string {
  const { section } = extractSection(xml);
  const elements = collectDocElements(section);
  let pi = 0;
  for (const el of elements) {
    if (el.type === 'p') {
      if (pi === paraIdx) return compactParagraph(el.content, pi);
      pi++;
    }
  }
  return 'Paragraph not found: p' + paraIdx;
}

function getTable(xml: string, tableIdx: number, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(section))) {
    if (ti === tableIdx) return compactTable(tm[0], tableIdx, listIdMap);
    ti++;
  }
  return 'Table not found: t' + tableIdx;
}

function getRow(xml: string, tableIdx: number, rowIdx: number, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(section))) {
    if (ti === tableIdx) {
      const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(tm[0]))) {
        if (ri === rowIdx) {
          // 해당 행만 간결하게
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

            const hasImage = cm[2].includes('<PICTURE');
            let content = '';
            if (hasImage) {
              content = '[IMAGE]';
            } else {
              content = extractInlineText(cm[2]);
            }

            out.push('  <CELL id="' + path + '"' + spanAttr + listId + '>' + content + '</CELL>');
          }
          out.push('</ROW>');
          return out.join('\n');
        }
        ri++;
      }
    }
    ti++;
  }
  return 'Row not found: t' + tableIdx + '.r' + rowIdx;
}

function getCell(xml: string, tableIdx: number, rowIdx: number, colIdx: number, listIdMap: Record<string, number>): string {
  const { section } = extractSection(xml);
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(section))) {
    if (ti === tableIdx) {
      const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(tm[0]))) {
        if (ri === rowIdx) {
          const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
          let cm;
          while ((cm = cellRe.exec(rm[0]))) {
            const ca = getAttr(cm[0], 'ColAddr');
            if (ca === colIdx) {
              const path = 't' + tableIdx + '.r' + rowIdx + '.c' + colIdx;
              const listId = listIdMap[path] !== undefined ? ' L=' + listIdMap[path] : '';
              const content = extractInlineText(cm[2]);
              return '<CELL id="' + path + '"' + listId + '>' + content + '</CELL>';
            }
          }
        }
        ri++;
      }
    }
    ti++;
  }
  return 'Cell not found';
}

function getRange(xml: string, start: ReturnType<typeof parseRange>['start'], end: NonNullable<ReturnType<typeof parseRange>['end']>, listIdMap: Record<string, number>): string {
  const startSeg = start.segments;
  const endSeg = end.segments;

  // p0~p2: 본문 문단 범위
  if (startSeg[0].type === 'p' && endSeg[0].type === 'p') {
    const from = startSeg[0].index;
    const to = endSeg[0].index;
    const { section } = extractSection(xml);
    const elements = collectDocElements(section);
    const out: string[] = [];
    let pi = 0;
    for (const el of elements) {
      if (el.type === 'p') {
        if (pi >= from && pi <= to) out.push(compactParagraph(el.content, pi));
        pi++;
      }
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
    const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
    let tm;
    let ti = 0;
    while ((tm = tableRe.exec(section))) {
      if (ti === tableIdx) {
        const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
        let rm;
        let ri = 0;
        while ((rm = rowRe.exec(tm[0]))) {
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
              const hasImage = cm[2].includes('<PICTURE');
              const content = hasImage ? '[IMAGE]' : extractInlineText(cm[2]);
              rowOut.push('    <CELL id="' + path + '"' + spanAttr + listId + '>' + content + '</CELL>');
            }
            rowOut.push('  </ROW>');
            out.push(rowOut.join('\n'));
          }
          ri++;
        }
      }
      ti++;
    }
    return out.join('\n');
  }

  return 'Range not supported: ' + JSON.stringify({ start: startSeg, end: endSeg });
}

/** XML에서 인라인 텍스트 추출 (CHAR + LINEBREAK) */
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

// ──────── styles — HEAD 섹션 스타일 조회 ────────

/** CharShape/ParaShape 조회 */
export function styles(xml: string, type?: string, id?: number): string {
  const headMatch = xml.match(/<HEAD[^>]*>([\s\S]*?)<\/HEAD>/);
  if (!headMatch) return 'HEAD section not found';

  const head = headMatch[1];
  const out: string[] = [];

  if (!type || type === 'CharShape') {
    out.push(formatStyleList(head, 'CharShape', id));
  }
  if (!type || type === 'ParaShape') {
    out.push(formatStyleList(head, 'ParaShape', id));
  }
  if (type && type !== 'CharShape' && type !== 'ParaShape') {
    out.push(formatStyleList(head, type, id));
  }

  return out.join('\n\n');
}

function formatStyleList(head: string, tagName: string, specificId?: number): string {
  const re = new RegExp('<' + tagName + '\\b([^>]*)/?>', 'g');
  const out: string[] = [];
  let m;
  let idx = 0;

  if (specificId !== undefined) {
    // 특정 ID 상세
    while ((m = re.exec(head))) {
      if (idx === specificId) {
        out.push(tagName + '[' + idx + ']:');
        const attrs = parseAllAttrs(m[1]);
        for (const [k, v] of Object.entries(attrs)) {
          out.push('  ' + k + ': ' + v);
        }
        return out.join('\n');
      }
      idx++;
    }
    return tagName + '[' + specificId + ']: not found';
  }

  // 전체 목록 (요약)
  out.push(tagName + ':');
  while ((m = re.exec(head))) {
    const attrs = parseAllAttrs(m[1]);
    const summary = summarizeStyle(tagName, attrs);
    out.push('  ' + idx + ': ' + summary);
    idx++;
  }

  if (idx === 0) out.push('  (none)');
  return out.join('\n');
}

function summarizeStyle(tagName: string, attrs: Record<string, string>): string {
  if (tagName === 'CharShape') {
    const parts: string[] = [];
    if (attrs['Height']) parts.push(Math.round(parseInt(attrs['Height']) / 100) + 'pt');
    if (attrs['Bold'] === '1') parts.push('볼드');
    if (attrs['Italic'] === '1') parts.push('이탤릭');
    if (attrs['Underline'] && attrs['Underline'] !== 'None') parts.push('밑줄');
    if (attrs['TextColor'] && attrs['TextColor'] !== '0') parts.push('색:' + attrs['TextColor']);
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
