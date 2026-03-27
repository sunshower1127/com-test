/**
 * XML 정제 — 화이트리스트 기반 속성 필터링, 간결/상세 포맷팅
 * 순수 함수: XML 입력 → 문자열 출력
 */

import { ATTR_WHITELIST } from './types';
import { getAttr, unescapeXml } from './path-resolver';

// ──────── 간결 (compact) 포맷 ────────

/** 본문 문단 → 간결 한 줄 */
export function compactParagraph(pXml: string, pIdx: number): string {
  const content = extractInlineContent(pXml);
  return '<P id="p' + pIdx + '">' + content + '</P>';
}

/** 표 → 간결 XML (스타일 제거, 텍스트만) */
export function compactTable(tableXml: string, tIdx: number, listIdMap: Record<string, number>): string {
  const out: string[] = [];

  const rowCountMatch = tableXml.match(/RowCount="(\d+)"/);
  const colCountMatch = tableXml.match(/ColCount="(\d+)"/);
  const rows = rowCountMatch ? rowCountMatch[1] : '?';
  const cols = colCountMatch ? colCountMatch[1] : '?';

  out.push('<TABLE id="t' + tIdx + '" rows=' + rows + ' cols=' + cols + '>');

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
      const path = 't' + tIdx + '.r' + ri + '.c' + colAddr;
      const listId = listIdMap[path] !== undefined ? ' L=' + listIdMap[path] : '';
      const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';

      const cellInner = cm[2];
      const hasImage = cellInner.includes('<PICTURE');

      // P 태그 수 확인
      const pMatches: string[] = [];
      const pRe = /<P[\s\S]*?<\/P>/g;
      let pm;
      while ((pm = pRe.exec(cellInner))) pMatches.push(pm[0]);

      let content = '';
      if (hasImage) {
        content = '[IMAGE]';
      } else if (pMatches.length <= 1) {
        content = extractInlineContent(cellInner);
      } else {
        // P 여러 개: P 유지
        const pLines = pMatches.map(pxml => '\n      <P>' + extractInlineContent(pxml) + '</P>');
        content = pLines.join('') + '\n    ';
      }

      out.push('    <CELL id="' + path + '"' + spanAttr + listId + '>' + content + '</CELL>');
    }

    out.push('  </ROW>');
    ri++;
  }

  out.push('</TABLE>');
  return out.join('\n');
}

/** XML 내의 CHAR/LINEBREAK에서 텍스트만 추출 */
function extractInlineContent(xml: string): string {
  const parts: string[] = [];
  const re = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<LINEBREAK\s*\/?>/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('<LINEBREAK')) {
      parts.push('\\n');
    } else if (m[1]) {
      const t = unescapeXml(m[1]).replace(/\n/g, '\\n').replace(/\r/g, '');
      parts.push(t);
    }
  }
  return parts.join('');
}

// ──────── 상세 (detail) 포맷 ────────

/** 표 → 상세 XML (화이트리스트 속성 유지) */
export function cleanTable(tableXml: string, tableIdx: number, baseIndent: number, listIdMap: Record<string, number>): string {
  const indent = '  '.repeat(baseIndent);
  const out: string[] = [];

  const tableTag = tableXml.match(/<TABLE([^>]*)>/);
  const tableAttrs = tableTag ? filterAttrs('TABLE', tableTag[1]) : '';
  out.push(indent + '<TABLE id="t' + tableIdx + '"' + tableAttrs + '>');

  const rowRe = /<ROW[^>]*>([\s\S]*?)<\/ROW>/g;
  let rm;
  let ri = 0;
  while ((rm = rowRe.exec(tableXml))) {
    out.push(indent + '  <ROW>');

    const cellRe = /<CELL\b([^>]*?)>([\s\S]*?)<\/CELL>/g;
    let cm;
    while ((cm = cellRe.exec(rm[0]))) {
      const cellAttrs = filterAttrs('CELL', cm[1]);
      const colAddr = getAttr(cm[0], 'ColAddr');
      const path = 't' + tableIdx + '.r' + ri + '.c' + colAddr;
      const listId = listIdMap[path] !== undefined ? ' L="' + listIdMap[path] + '"' : '';
      out.push(indent + '    <CELL id="' + path + '"' + cellAttrs + listId + '>');

      const paralistMatch = cm[2].match(/<PARALIST([^>]*)>([\s\S]*?)<\/PARALIST>/);
      if (paralistMatch) {
        const plAttrs = filterAttrs('PARALIST', paralistMatch[1]);
        if (plAttrs.trim()) out.push(indent + '      <PARALIST' + plAttrs + '>');

        const pInCellRe = /<P[\s\S]*?<\/P>/g;
        let pcm;
        while ((pcm = pInCellRe.exec(paralistMatch[2]))) {
          out.push(cleanParagraph(pcm[0], undefined, baseIndent + 3));
        }

        if (plAttrs.trim()) out.push(indent + '      </PARALIST>');
      }

      out.push(indent + '    </CELL>');
    }

    out.push(indent + '  </ROW>');
    ri++;
  }

  // CAPTION 처리
  const captionMatch = tableXml.match(/<CAPTION([^>]*)>([\s\S]*?)<\/CAPTION>/);
  if (captionMatch) {
    const capAttrs = filterAttrs('CAPTION', captionMatch[1]);
    out.push(indent + '  <CAPTION' + capAttrs + '>');
    const pInCapRe = /<P[\s\S]*?<\/P>/g;
    let pcap;
    while ((pcap = pInCapRe.exec(captionMatch[2]))) {
      out.push(cleanParagraph(pcap[0], undefined, baseIndent + 2));
    }
    out.push(indent + '  </CAPTION>');
  }

  out.push(indent + '</TABLE>');
  return out.join('\n');
}

/** 문단 → 상세 XML */
export function cleanParagraph(pXml: string, id: string | undefined, baseIndent: number): string {
  const indent = '  '.repeat(baseIndent);
  const out: string[] = [];

  const pTag = pXml.match(/<P([^>]*)>/);
  const pAttrs = pTag ? filterAttrs('P', pTag[1]) : '';
  const idAttr = id ? ' id="' + id + '"' : '';
  out.push(indent + '<P' + idAttr + pAttrs + '>');

  const textRe = /<TEXT\b([^>]*?)>([\s\S]*?)<\/TEXT>|<TEXT\b([^/]*?)\/>/g;
  let txm;
  while ((txm = textRe.exec(pXml))) {
    if (txm[2] !== undefined) {
      const textAttrs = filterAttrs('TEXT', txm[1]);
      const inner = txm[2];

      if (inner.includes('<PICTURE')) {
        out.push(indent + '  <TEXT' + textAttrs + '>');
        out.push(cleanPicture(inner, baseIndent + 2));
        out.push(indent + '  </TEXT>');
      } else {
        const content = cleanTextContent(inner);
        if (content) {
          out.push(indent + '  <TEXT' + textAttrs + '>' + content + '</TEXT>');
        } else {
          out.push(indent + '  <TEXT' + textAttrs + '/>');
        }
      }
    } else {
      const textAttrs = filterAttrs('TEXT', txm[3] || '');
      out.push(indent + '  <TEXT' + textAttrs + '/>');
    }
  }

  out.push(indent + '</P>');
  return out.join('\n');
}

/** TEXT 내부 → CHAR + LINEBREAK 정제 */
function cleanTextContent(inner: string): string {
  const parts: string[] = [];
  const re = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<CHAR\b[^/]*?\/>|<LINEBREAK\/?>|<LINEBREAK\s*\/>/g;
  let m;
  while ((m = re.exec(inner))) {
    if (m[0].startsWith('<LINEBREAK')) {
      parts.push('<LINEBREAK/>');
    } else if (m[1] !== undefined) {
      const text = unescapeXml(m[1]).replace(/\n/g, '\\n').replace(/\r/g, '');
      parts.push('<CHAR>' + text + '</CHAR>');
    }
  }
  return parts.join('');
}

/** PICTURE 정제 */
function cleanPicture(inner: string, baseIndent: number): string {
  const indent = '  '.repeat(baseIndent);
  const out: string[] = [];

  const picMatch = inner.match(/<PICTURE([^>]*)>([\s\S]*?)<\/PICTURE>/);
  if (!picMatch) return indent + '<!-- PICTURE not parsed -->';

  const picAttrs = filterAttrs('PICTURE', picMatch[1]);
  out.push(indent + '<PICTURE' + picAttrs + '>');

  const soMatch = picMatch[2].match(/<SHAPEOBJECT([^>]*)>/);
  if (soMatch) {
    const soAttrs = filterAttrs('SHAPEOBJECT', soMatch[1]);
    out.push(indent + '  <SHAPEOBJECT' + soAttrs + '>');

    const sizeMatch = picMatch[2].match(/<SIZE([^/]*?)\/>/);
    if (sizeMatch) out.push(indent + '    <SIZE' + filterAttrs('SIZE', sizeMatch[1]) + '/>');

    const posMatch = picMatch[2].match(/<POSITION([^/]*?)\/>/);
    if (posMatch) out.push(indent + '    <POSITION' + filterAttrs('POSITION', posMatch[1]) + '/>');

    out.push(indent + '  </SHAPEOBJECT>');
  }

  const imgMatch = picMatch[2].match(/<IMAGE([^/]*?)\/>/);
  if (imgMatch) out.push(indent + '  <IMAGE' + filterAttrs('IMAGE', imgMatch[1]) + '/>');

  out.push(indent + '</PICTURE>');
  return out.join('\n');
}

/** 화이트리스트 기반 속성 필터 */
export function filterAttrs(tagName: string, attrStr: string): string {
  const whitelist = ATTR_WHITELIST[tagName];
  if (!whitelist) return '';
  if (whitelist === '*') return attrStr;
  if (whitelist.length === 0) return '';

  const result: string[] = [];
  const re = /(\w+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(attrStr))) {
    if ((whitelist as string[]).includes(m[1])) {
      result.push(m[1] + '="' + m[2] + '"');
    }
  }
  return result.length > 0 ? ' ' + result.join(' ') : '';
}
