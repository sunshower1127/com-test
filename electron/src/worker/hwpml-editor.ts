/**
 * HWPML2X 편집기 — set/commit 패턴으로 XML 부분 수정
 *
 * 사용법 (sandbox에서):
 *   getXml()                        // 문서에서 XML 로드
 *   structure()                     // 구조맵 출력
 *   set("t0.r1.c3", "홍길동")       // 표0 행1 열3 전체 텍스트
 *   set("t0.r1.c3.p0", "첫줄만")   // 표0 행1 열3 첫 문단만
 *   set("p2", "본문 세번째 문단")   // 본문 문단2 전체
 *   commit()                        // 수정사항 문서에 반영
 */

type ComHandle = unknown;

interface ComBridge {
  comGet(handle: ComHandle, prop: string): unknown;
  comPut(handle: ComHandle, prop: string, value: unknown): void;
  comCallWith(handle: ComHandle, method: string, args: unknown[]): unknown;
}

let _bridge: ComBridge;
let _hwpHandle: ComHandle;
let _xml: string = '';
let _pendingChanges: Array<{ path: string; value: string }> = [];
let _listIdMap: Record<string, number> = {};  // path → List ID 캐시

export function initHwpmlEditor(bridge: ComBridge) {
  _bridge = bridge;
}

export function setHwpHandle(handle: ComHandle) {
  _hwpHandle = handle;
}

/** 현재 문서에서 HWPML2X 로드 */
export function getXml(): string {
  const proxy = _hwpHandle;
  _xml = String(_bridge.comCallWith(proxy, 'GetTextFile', ['HWPML2X', '']));
  _pendingChanges = [];
  return 'XML loaded (' + _xml.length + ' chars)';
}

/** 변경 등록 */
export function set(path: string, value: string): string {
  _pendingChanges.push({ path, value });
  return 'queued: ' + path + ' = "' + (value.length > 30 ? value.substring(0, 30) + '...' : value) + '"';
}

/** 등록된 변경사항을 XML에 적용하고 문서에 반영 */
export function commit(): string {
  if (!_xml) {
    throw new Error('XML not loaded. Call getXml() first.');
  }
  if (_pendingChanges.length === 0) {
    return 'No pending changes.';
  }

  let xml = _xml;

  for (const change of _pendingChanges) {
    xml = applyChange(xml, change.path, change.value);
  }

  // SelectAll → Delete → SetTextFile
  _bridge.comCallWith(_hwpHandle, 'Run', ['SelectAll']);
  _bridge.comCallWith(_hwpHandle, 'Run', ['Delete']);
  _bridge.comCallWith(_hwpHandle, 'SetTextFile', [xml, 'HWPML2X', '']);

  const count = _pendingChanges.length;
  _xml = xml;
  _pendingChanges = [];
  return 'Committed ' + count + ' change(s).';
}

/** 셀 텍스트 끝에 텍스트 추가 (원본 보존) */
export function append(path: string, text: string): string {
  _pendingChanges.push({ path, value: '{{APPEND}}' + text });
  return 'queued append: ' + path + ' += "' + text + '"';
}

/** 모든 표 셀에 마커를 append → RepeatFind로 List ID 매핑 → 마커 제거 */
export function mapListIds(): string {
  if (!_xml) throw new Error('XML not loaded. Call xml.load() first.');

  const tables = parseTables(_xml);
  const markerPrefix = '{{#LID_';
  const paths: string[] = [];

  // 1. 각 셀에 유니크 마커 append
  let xml = _xml;
  for (let ti = 0; ti < tables.length; ti++) {
    tables[ti].forEach((row, ri) => {
      row.forEach((cell) => {
        const path = 't' + ti + '.r' + ri + '.c' + cell.col;
        const marker = markerPrefix + path + '}}';
        paths.push(path);
        // appendCellText: 마지막 CHAR 태그 안에 마커 추가
        xml = appendToCellXml(xml, ['t' + ti, 'r' + ri, 'c' + cell.col], marker);
      });
    });
  }

  // 2. 문서에 반영
  _bridge.comCallWith(_hwpHandle, 'Run', ['SelectAll']);
  _bridge.comCallWith(_hwpHandle, 'Run', ['Delete']);
  _bridge.comCallWith(_hwpHandle, 'SetTextFile', [xml, 'HWPML2X', '']);

  // 3. 각 마커를 RepeatFind로 찾아서 List ID 기록
  _listIdMap = {};
  _bridge.comCallWith(_hwpHandle, 'Run', ['MoveDocBegin']);

  for (const path of paths) {
    const marker = markerPrefix + path + '}}';

    // HParameterSet.HFindReplace 설정
    const hwp = _hwpHandle;
    const hAction = _bridge.comGet(hwp, 'HAction');
    const hParamSet = _bridge.comGet(hwp, 'HParameterSet');
    const hFindReplace = _bridge.comGet(hParamSet, 'HFindReplace');
    const hSet = _bridge.comGet(hFindReplace, 'HSet');

    _bridge.comCallWith(hAction, 'GetDefault', ['RepeatFind', hSet]);
    _bridge.comPut(hFindReplace, 'FindString', marker);
    _bridge.comPut(hFindReplace, 'IgnoreMessage', 1);

    const found = _bridge.comCallWith(hAction, 'Execute', ['RepeatFind', hSet]);
    if (found) {
      const ps = _bridge.comCallWith(hwp, 'GetPosBySet', []);
      const listId = _bridge.comCallWith(ps, 'Item', ['List']);
      if (typeof listId === 'number') {
        _listIdMap[path] = listId;
      }
    }
  }

  // 4. 마커 전체 제거 — HWPML2X에서 직접 제거 후 재로드
  let cleanXml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
  // {{#LID_...}} 패턴을 모두 제거 (XML 이스케이프된 버전도 처리)
  cleanXml = cleanXml.replace(/\{\{#LID_[^}]*\}\}/g, '');
  cleanXml = cleanXml.replace(/\{\{#LID_[^}]*}}|%7B%7B#LID_[^%]*%7D%7D/g, '');

  _bridge.comCallWith(_hwpHandle, 'Run', ['SelectAll']);
  _bridge.comCallWith(_hwpHandle, 'Run', ['Delete']);
  _bridge.comCallWith(_hwpHandle, 'SetTextFile', [cleanXml, 'HWPML2X', '']);

  // 5. XML 다시 로드
  _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));

  const mapped = Object.keys(_listIdMap).length;
  return 'Mapped ' + mapped + '/' + paths.length + ' cells. List IDs cached.';
}

function appendToCellXml(xml: string, parts: string[], text: string): string {
  const tableIdx = parseInt(parts[0].substring(1));
  const rowIdx = parseInt(parts[1].substring(1));
  const colIdx = parseInt(parts[2].substring(1));

  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(xml))) {
    if (ti === tableIdx) {
      const tableStart = tm.index;
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
              const cellXml = cm[0];
              // 마지막 </CHAR> 앞에 텍스트 삽입. 없으면 </TEXT> 앞에 <CHAR>텍스트</CHAR> 삽입
              let newCell: string;
              const lastCharClose = cellXml.lastIndexOf('</CHAR>');
              if (lastCharClose >= 0) {
                newCell = cellXml.substring(0, lastCharClose) + escapeXml(text) + cellXml.substring(lastCharClose);
              } else {
                // CHAR가 없는 경우 (빈 TEXT 또는 셀프클로징)
                const lastTextClose = cellXml.lastIndexOf('</TEXT>');
                if (lastTextClose >= 0) {
                  newCell = cellXml.substring(0, lastTextClose) + '<CHAR>' + escapeXml(text) + '</CHAR>' + cellXml.substring(lastTextClose);
                } else {
                  // 셀프클로징 TEXT: <TEXT .../> → <TEXT ...><CHAR>텍스트</CHAR></TEXT>
                  const selfTextRe = /<TEXT([^/]*?)\/>/;
                  const stm = selfTextRe.exec(cellXml);
                  if (stm) {
                    const replacement = '<TEXT' + stm[1] + '><CHAR>' + escapeXml(text) + '</CHAR></TEXT>';
                    newCell = cellXml.substring(0, stm.index) + replacement + cellXml.substring(stm.index + stm[0].length);
                  } else {
                    newCell = cellXml; // 변경 불가
                  }
                }
              }
              const absStart = tableStart + rm.index + cm.index;
              const absEnd = absStart + cellXml.length;
              return xml.substring(0, absStart) + newCell + xml.substring(absEnd);
            }
          }
          return xml; // cell not found, skip
        }
        ri++;
      }
      return xml;
    }
    ti++;
  }
  return xml;
}

/** 특정 경로의 원본 XML 반환 */
export function raw(path: string): string {
  if (!_xml) throw new Error('XML not loaded. Call xml.load() first.');

  const parts = path.split('.');

  if (parts[0].startsWith('t')) {
    return findCellXml(_xml, parts);
  } else if (parts[0].startsWith('p')) {
    return findParaXml(_xml, parts);
  }

  throw new Error('Unknown path: ' + path);
}

/** 특정 경로의 XML을 교체하고 문서에 반영 */
export function rawSet(path: string, newXml: string): string {
  if (!_xml) throw new Error('XML not loaded. Call xml.load() first.');

  // XML 기본 검증: 태그 균형 체크
  const openTags = (newXml.match(/<[a-zA-Z][^/>]*>/g) || []).length;
  const closeTags = (newXml.match(/<\/[a-zA-Z][^>]*>/g) || []).length;
  const selfClose = (newXml.match(/<[a-zA-Z][^>]*\/>/g) || []).length;
  if (openTags !== closeTags + selfClose && openTags !== closeTags) {
    // 간단한 휴리스틱: 열린 태그 수 ≠ 닫힌 태그 수 + 셀프클로징이면 경고
    const msg = 'XML 태그 불균형 경고: open=' + openTags + ' close=' + closeTags + ' selfClose=' + selfClose;
    // 경고만 하고 계속 진행 (완벽한 검증은 아니므로)
    console.warn(msg);
  }

  const parts = path.split('.');
  let xml = _xml;

  if (parts[0].startsWith('t')) {
    xml = replaceCellRaw(_xml, parts, newXml);
  } else if (parts[0].startsWith('p')) {
    xml = replaceParaRaw(_xml, parts, newXml);
  } else {
    throw new Error('Unknown path: ' + path);
  }

  // 문서에 반영
  _bridge.comCallWith(_hwpHandle, 'Run', ['SelectAll']);
  _bridge.comCallWith(_hwpHandle, 'Run', ['Delete']);
  _bridge.comCallWith(_hwpHandle, 'SetTextFile', [xml, 'HWPML2X', '']);

  _xml = xml;
  return 'SET ' + path + ' applied.';
}

function findCellXml(xml: string, parts: string[]): string {
  const tableIdx = parseInt(parts[0].substring(1));
  const rowIdx = parseInt(parts[1].substring(1));
  const colIdx = parseInt(parts[2].substring(1));

  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(xml))) {
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
            if (ca === colIdx) return cm[0];
          }
          throw new Error('Cell not found: c' + colIdx);
        }
        ri++;
      }
      throw new Error('Row not found: r' + rowIdx);
    }
    ti++;
  }
  throw new Error('Table not found: t' + tableIdx);
}

function findParaXml(xml: string, parts: string[]): string {
  const paraIdx = parseInt(parts[0].substring(1));

  const sectionMatch = xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!sectionMatch) throw new Error('SECTION not found');

  const section = sectionMatch[1];
  const noTables = section.replace(/<TABLE[^>]*>[\s\S]*?<\/TABLE>/g, '{{TABLE}}');
  const pRe = /<P[\s\S]*?<\/P>/g;
  let m;
  let pi = 0;
  while ((m = pRe.exec(noTables))) {
    if (m[0].includes('{{TABLE}}')) continue;
    if (pi === paraIdx) return m[0];
    pi++;
  }
  throw new Error('Paragraph not found: p' + paraIdx);
}

function replaceCellRaw(xml: string, parts: string[], newCellXml: string): string {
  const tableIdx = parseInt(parts[0].substring(1));
  const rowIdx = parseInt(parts[1].substring(1));
  const colIdx = parseInt(parts[2].substring(1));

  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(xml))) {
    if (ti === tableIdx) {
      const tableStart = tm.index;
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
              const absStart = tableStart + rm.index + cm.index;
              const absEnd = absStart + cm[0].length;
              return xml.substring(0, absStart) + newCellXml + xml.substring(absEnd);
            }
          }
          throw new Error('Cell not found: c' + colIdx);
        }
        ri++;
      }
      throw new Error('Row not found: r' + rowIdx);
    }
    ti++;
  }
  throw new Error('Table not found: t' + tableIdx);
}

function replaceParaRaw(xml: string, parts: string[], newParaXml: string): string {
  const paraIdx = parseInt(parts[0].substring(1));

  const sectionMatch = xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!sectionMatch) throw new Error('SECTION not found');

  const section = sectionMatch[1];
  const sectionStart = sectionMatch.index! + sectionMatch[0].indexOf(section);

  const tableMatches: Array<{ start: number; end: number }> = [];
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  while ((tm = tableRe.exec(section))) {
    tableMatches.push({ start: tm.index, end: tm.index + tm[0].length });
  }

  const pRe = /<P[\s\S]*?<\/P>/g;
  let m;
  let pi = 0;
  while ((m = pRe.exec(section))) {
    const inTable = tableMatches.some(t => m!.index >= t.start && m!.index < t.end);
    if (inTable) continue;
    if (pi === paraIdx) {
      const absStart = sectionStart + m.index;
      const absEnd = absStart + m[0].length;
      return xml.substring(0, absStart) + newParaXml + xml.substring(absEnd);
    }
    pi++;
  }
  throw new Error('Paragraph not found: p' + paraIdx);
}

// ──────── Structure: 화이트리스트 기반 XML 정제 ────────

const ATTR_WHITELIST: Record<string, string[] | '*'> = {
  TABLE:     ['ColCount', 'RowCount', 'BorderFill'],
  ROW:       [],
  CELL:      ['ColAddr', 'RowAddr', 'ColSpan', 'RowSpan', 'Width', 'Height', 'BorderFill'],
  P:         ['ParaShape', 'Style'],
  TEXT:      ['CharShape'],
  CHAR:      [],
  LINEBREAK: [],
  PICTURE:   '*',
  SHAPEOBJECT: ['InstId'],
  SIZE:      ['Width', 'Height', 'WidthRelTo', 'HeightRelTo'],
  POSITION:  ['TreatAsChar', 'HorzAlign', 'VertAlign'],
  IMAGE:     '*',
  CAPTION:   ['Side'],
  PARALIST:  ['VertAlign'],
};

const REMOVE_TAGS = new Set([
  'CELLMARGIN', 'OUTSIDEMARGIN', 'INSIDEMARGIN',
  'RENDERINGINFO', 'ROTATIONINFO', 'SHAPECOMPONENT',
  'IMAGERECT', 'IMAGECLIP', 'EFFECTS', 'SHAPECOMMENT',
]);

/** 간결 구조맵 — 스타일 제거, 텍스트와 구조만 */
export function structure(): string {
  if (!_xml) {
    getXml();
  }

  const sectionMatch = _xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!sectionMatch) return 'SECTION not found';

  const section = sectionMatch[1];
  const elements = collectDocElements(section);
  const out: string[] = [];
  let pIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      out.push(compactParagraph(el.content, pIdx));
      pIdx++;
    } else {
      out.push(compactTable(el.content, tIdx));
      tIdx++;
    }
  }

  return out.join('\n');
}

function compactParagraph(pXml: string, pIdx: number): string {
  const texts: string[] = [];
  const re = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<LINEBREAK\s*\/?>/g;
  let m;
  while ((m = re.exec(pXml))) {
    if (m[0].startsWith('<LINEBREAK')) {
      texts.push('\\n');
    } else if (m[1]) {
      const t = m[1]
        .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/\n/g, '\\n').replace(/\r/g, '');
      texts.push(t);
    }
  }
  const content = texts.join('');
  return '<P id="p' + pIdx + '">' + content + '</P>';
}

function compactTable(tableXml: string, tIdx: number): string {
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
      const listId = _listIdMap[path] !== undefined ? ' L=' + _listIdMap[path] : '';

      const spanAttr = (colSpan > 1 || rowSpan > 1) ? ' span="' + colSpan + 'x' + rowSpan + '"' : '';

      // 셀 내용: 이미지 / 텍스트
      const cellInner = cm[2];
      const hasImage = cellInner.includes('<PICTURE');

      // P 태그 수 확인
      const pMatches: string[] = [];
      const pRe = /<P[\s\S]*?<\/P>/g;
      let pm;
      while ((pm = pRe.exec(cellInner))) {
        pMatches.push(pm[0]);
      }

      let content = '';
      if (hasImage) {
        content = '[IMAGE]';
      } else if (pMatches.length <= 1) {
        // P 1개: 뭉개고 텍스트만
        const texts: string[] = [];
        const charRe = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<LINEBREAK\s*\/?>/g;
        let chm;
        while ((chm = charRe.exec(cellInner))) {
          if (chm[0].startsWith('<LINEBREAK')) {
            texts.push('\\n');
          } else if (chm[1]) {
            const t = chm[1]
              .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
              .replace(/&gt;/g, '>').replace(/&quot;/g, '"')
              .replace(/\n/g, '\\n').replace(/\r/g, '');
            texts.push(t);
          }
        }
        content = texts.join('');
      } else {
        // P 여러 개: P 유지
        const pLines: string[] = [];
        pMatches.forEach((pxml) => {
          const texts: string[] = [];
          const charRe = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<LINEBREAK\s*\/?>/g;
          let chm;
          while ((chm = charRe.exec(pxml))) {
            if (chm[0].startsWith('<LINEBREAK')) {
              texts.push('\\n');
            } else if (chm[1]) {
              const t = chm[1]
                .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>').replace(/&quot;/g, '"')
                .replace(/\n/g, '\\n').replace(/\r/g, '');
              texts.push(t);
            }
          }
          pLines.push('\n      <P>' + texts.join('') + '</P>');
        });
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

/** 문서 내 P(본문) + TABLE을 순서대로 수집 */
function collectDocElements(section: string): Array<{ type: 'p' | 'table'; index: number; content: string }> {
  const elements: Array<{ type: 'p' | 'table'; index: number; content: string }> = [];

  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  while ((tm = tableRe.exec(section))) {
    elements.push({ type: 'table', index: tm.index, content: tm[0] });
  }

  const marker = '{{TABLE}}';
  const tableMatches: Array<{ start: number; length: number }> = [];
  const tableRe2 = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm2;
  while ((tm2 = tableRe2.exec(section))) {
    tableMatches.push({ start: tm2.index, length: tm2[0].length });
  }

  const noTables = section.replace(/<TABLE[^>]*>[\s\S]*?<\/TABLE>/g, marker);
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

/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 정제 XML */
export function structureDetail(): string {
  if (!_xml) {
    getXml();
  }

  const sectionMatch = _xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!sectionMatch) return 'SECTION not found';

  const section = sectionMatch[1];
  const elements = collectDocElements(section);

  // 순서대로 정제 XML 출력
  const out: string[] = [];
  let pIdx = 0;
  let tIdx = 0;

  for (const el of elements) {
    if (el.type === 'p') {
      out.push(cleanParagraph(el.content, 'p' + pIdx, 0));
      pIdx++;
    } else {
      out.push(cleanTable(el.content, tIdx, 0));
      tIdx++;
    }
  }

  return out.join('\n');
}

function cleanTable(tableXml: string, tableIdx: number, baseIndent: number): string {
  const indent = '  '.repeat(baseIndent);
  const out: string[] = [];

  // TABLE 속성 정제
  const tableTag = tableXml.match(/<TABLE([^>]*)>/);
  const tableAttrs = tableTag ? filterAttrs('TABLE', tableTag[1]) : '';
  const listIdAttr = _listIdMap['t' + tableIdx] !== undefined ? ' L="' + _listIdMap['t' + tableIdx] + '"' : '';
  out.push(indent + '<TABLE id="t' + tableIdx + '"' + tableAttrs + listIdAttr + '>');

  // ROW/CELL 순회
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
      const listId = _listIdMap[path] !== undefined ? ' L="' + _listIdMap[path] + '"' : '';
      out.push(indent + '    <CELL id="' + path + '"' + cellAttrs + listId + '>');

      // CELL 내부: PARALIST > P
      const paralistMatch = cm[2].match(/<PARALIST([^>]*)>([\s\S]*?)<\/PARALIST>/);
      if (paralistMatch) {
        const plAttrs = filterAttrs('PARALIST', paralistMatch[1]);
        // VertAlign만 기본값 아닌 경우 표시
        if (plAttrs.trim()) {
          out.push(indent + '      <PARALIST' + plAttrs + '>');
        }

        const pInCellRe = /<P[\s\S]*?<\/P>/g;
        let pcm;
        while ((pcm = pInCellRe.exec(paralistMatch[2]))) {
          out.push(cleanParagraph(pcm[0], undefined, baseIndent + 3));
        }

        if (plAttrs.trim()) {
          out.push(indent + '      </PARALIST>');
        }
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

function cleanParagraph(pXml: string, id: string | undefined, baseIndent: number): string {
  const indent = '  '.repeat(baseIndent);
  const out: string[] = [];

  // P 속성
  const pTag = pXml.match(/<P([^>]*)>/);
  const pAttrs = pTag ? filterAttrs('P', pTag[1]) : '';
  const idAttr = id ? ' id="' + id + '"' : '';
  out.push(indent + '<P' + idAttr + pAttrs + '>');

  // TEXT 태그 순회 (열린 태그 + 셀프클로징)
  const textRe = /<TEXT\b([^>]*?)>([\s\S]*?)<\/TEXT>|<TEXT\b([^/]*?)\/>/g;
  let txm;
  while ((txm = textRe.exec(pXml))) {
    if (txm[2] !== undefined) {
      // 열린 TEXT
      const textAttrs = filterAttrs('TEXT', txm[1]);
      const inner = txm[2];

      // PICTURE 포함 여부
      if (inner.includes('<PICTURE')) {
        out.push(indent + '  <TEXT' + textAttrs + '>');
        out.push(cleanPicture(inner, baseIndent + 2));
        out.push(indent + '  </TEXT>');
      } else {
        // CHAR + LINEBREAK
        const content = cleanTextContent(inner);
        if (content) {
          out.push(indent + '  <TEXT' + textAttrs + '>' + content + '</TEXT>');
        } else {
          out.push(indent + '  <TEXT' + textAttrs + '/>');
        }
      }
    } else {
      // 셀프클로징 TEXT
      const textAttrs = filterAttrs('TEXT', txm[3] || '');
      out.push(indent + '  <TEXT' + textAttrs + '/>');
    }
  }

  out.push(indent + '</P>');
  return out.join('\n');
}

function cleanTextContent(inner: string): string {
  const parts: string[] = [];
  // CHAR 내용 + LINEBREAK 순서대로
  const re = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>|<CHAR\b[^/]*?\/>|<LINEBREAK\/?>|<LINEBREAK\s*\/>/g;
  let m;
  while ((m = re.exec(inner))) {
    if (m[0].startsWith('<LINEBREAK')) {
      parts.push('<LINEBREAK/>');
    } else if (m[1] !== undefined) {
      const text = m[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"');
      parts.push('<CHAR>' + text.replace(/\n/g, '\\n').replace(/\r/g, '') + '</CHAR>');
    }
    // 셀프클로징 CHAR는 무시 (빈 커서 마커)
  }
  return parts.join('');
}

function cleanPicture(inner: string, baseIndent: number): string {
  const indent = '  '.repeat(baseIndent);
  const out: string[] = [];

  // PICTURE 태그
  const picMatch = inner.match(/<PICTURE([^>]*)>([\s\S]*?)<\/PICTURE>/);
  if (!picMatch) return indent + '<!-- PICTURE not parsed -->';

  const picAttrs = filterAttrs('PICTURE', picMatch[1]);
  out.push(indent + '<PICTURE' + picAttrs + '>');

  // SHAPEOBJECT > SIZE
  const soMatch = picMatch[2].match(/<SHAPEOBJECT([^>]*)>/);
  if (soMatch) {
    const soAttrs = filterAttrs('SHAPEOBJECT', soMatch[1]);
    out.push(indent + '  <SHAPEOBJECT' + soAttrs + '>');

    const sizeMatch = picMatch[2].match(/<SIZE([^/]*?)\/>/);
    if (sizeMatch) {
      const sizeAttrs = filterAttrs('SIZE', sizeMatch[1]);
      out.push(indent + '    <SIZE' + sizeAttrs + '/>');
    }

    const posMatch = picMatch[2].match(/<POSITION([^/]*?)\/>/);
    if (posMatch) {
      const posAttrs = filterAttrs('POSITION', posMatch[1]);
      out.push(indent + '    <POSITION' + posAttrs + '/>');
    }

    out.push(indent + '  </SHAPEOBJECT>');
  }

  // IMAGE
  const imgMatch = picMatch[2].match(/<IMAGE([^/]*?)\/>/);
  if (imgMatch) {
    const imgAttrs = filterAttrs('IMAGE', imgMatch[1]);
    out.push(indent + '  <IMAGE' + imgAttrs + '/>');
  }

  out.push(indent + '</PICTURE>');
  return out.join('\n');
}

function filterAttrs(tagName: string, attrStr: string): string {
  const whitelist = ATTR_WHITELIST[tagName];
  if (!whitelist) return '';  // 알 수 없는 태그 → 속성 전부 제거
  if (whitelist === '*') return attrStr;  // 전부 유지

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

// ──────── Internal ────────

interface CellInfo {
  col: number;
  row: number;
  colSpan: number;
  rowSpan: number;
  text: string;
}

function parseTables(xml: string): CellInfo[][][] {
  const tables: CellInfo[][][] = [];
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
  const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;

  let tm;
  while ((tm = tableRe.exec(xml))) {
    const rows: CellInfo[][] = [];
    let rm;
    rowRe.lastIndex = 0;
    while ((rm = rowRe.exec(tm[0]))) {
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
    tables.push(rows);
  }
  return tables;
}

function getAttr(tag: string, name: string): number | null {
  const re = new RegExp(name + '="(\\d+)"');
  const m = re.exec(tag);
  return m ? +m[1] : null;
}

function extractText(cellXml: string): string {
  const texts: string[] = [];
  // CHAR 태그 안의 텍스트 추출 (엔티티 포함)
  const charRe = /<CHAR\b[^>]*>([\s\S]*?)<\/CHAR>/g;
  let m;
  while ((m = charRe.exec(cellXml))) {
    const t = m[1]
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .trim();
    if (t) texts.push(t);
  }
  return texts.join(' ');
}

function extractBodyParagraphs(xml: string): string[] {
  // BODY > SECTION 내의 P 태그 중 TABLE 밖에 있는 것들
  const sectionMatch = xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!sectionMatch) return [];

  const section = sectionMatch[1];
  const paragraphs: string[] = [];

  // TABLE을 제거하고 남은 P 태그들이 본문 문단
  const noTables = section.replace(/<TABLE[^>]*>[\s\S]*?<\/TABLE>/g, '{{TABLE}}');
  const pRe = /<P[\s\S]*?<\/P>/g;
  let m;
  while ((m = pRe.exec(noTables))) {
    if (m[0].includes('{{TABLE}}')) continue;
    const texts: string[] = [];
    const charRe = /<CHAR[^>]*>([^<]*)<\/CHAR>/g;
    let cm;
    while ((cm = charRe.exec(m[0]))) {
      texts.push(cm[1]);
    }
    paragraphs.push(texts.join(''));
  }

  return paragraphs;
}

function applyChange(xml: string, path: string, value: string): string {
  const parts = path.split('.');

  if (parts[0].startsWith('t')) {
    // 표 셀 수정
    return applyTableChange(xml, parts, value);
  } else if (parts[0].startsWith('p')) {
    // 본문 문단 수정
    return applyParaChange(xml, parts, value);
  }

  throw new Error('Unknown path: ' + path);
}

function applyTableChange(xml: string, parts: string[], value: string): string {
  const tableIdx = parseInt(parts[0].substring(1));
  const rowIdx = parseInt(parts[1].substring(1));
  const colIdx = parseInt(parts[2].substring(1));
  const paraIdx = parts.length > 3 ? parseInt(parts[3].substring(1)) : -1; // -1 = 전체

  // n번째 TABLE 찾기
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  let ti = 0;
  while ((tm = tableRe.exec(xml))) {
    if (ti === tableIdx) {
      const tableStart = tm.index;
      const tableEnd = tm.index + tm[0].length;
      const tableXml = tm[0];

      // ROW에서 rowIdx번째 찾기
      const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
      let rm;
      let ri = 0;
      while ((rm = rowRe.exec(tableXml))) {
        if (ri === rowIdx) {
          // CELL에서 ColAddr=colIdx 찾기
          const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
          let cm: RegExpExecArray | null = null;
          let cellMatch: RegExpExecArray | null = null;
          while ((cellMatch = cellRe.exec(rm[0]))) {
            const ca = getAttr(cellMatch[0], 'ColAddr');
            if (ca === colIdx) {
              cm = cellMatch;
              break;
            }
          }
          if (!cm) {
            throw new Error('Cell not found: t' + tableIdx + '.r' + rowIdx + '.c' + colIdx);
          }

          const cellXml = cm[0];
          const newCellXml = replaceCellText(cellXml, value, paraIdx);

          // 원본 XML에서 해당 셀만 교체
          const cellAbsStart = tableStart + rm.index + cm.index;
          const cellAbsEnd = cellAbsStart + cellXml.length;
          return xml.substring(0, cellAbsStart) + newCellXml + xml.substring(cellAbsEnd);
        }
        ri++;
      }
      throw new Error('Row not found: r' + rowIdx);
    }
    ti++;
  }
  throw new Error('Table not found: t' + tableIdx);
}

function applyParaChange(xml: string, parts: string[], value: string): string {
  const paraIdx = parseInt(parts[0].substring(1));

  const sectionMatch = xml.match(/<SECTION[^>]*>([\s\S]*?)<\/SECTION>/);
  if (!sectionMatch) throw new Error('SECTION not found');

  const section = sectionMatch[1];

  // extractBodyParagraphs와 동일한 방식: TABLE을 마커로 치환 후 P 찾기
  // 단, 원본 section에서의 위치를 역추적하기 위해 TABLE 위치 기록
  const tableMatches: Array<{ start: number; end: number; length: number }> = [];
  const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
  let tm;
  while ((tm = tableRe.exec(section))) {
    tableMatches.push({ start: tm.index, end: tm.index + tm[0].length, length: tm[0].length });
  }

  const marker = '{{TABLE}}';
  const noTables = section.replace(/<TABLE[^>]*>[\s\S]*?<\/TABLE>/g, marker);

  const pRe = /<P[\s\S]*?<\/P>/g;
  let pm;
  let pi = 0;
  while ((pm = pRe.exec(noTables))) {
    if (pm[0].includes(marker)) continue;

    if (pi === paraIdx) {
      const pXml = pm[0];
      const newPXml = replaceParaText(pXml, value);

      // noTables에서의 offset을 원본 section의 offset으로 변환
      let origOffset = pm.index;
      let markerIdx = 0;
      // noTables의 각 {{TABLE}} 위치 이전의 것들에 대해 길이 차이를 보정
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

      const sectionContentStart = xml.indexOf(section);
      const pAbsStart = sectionContentStart + origOffset;
      const pAbsEnd = pAbsStart + pXml.length;

      return xml.substring(0, pAbsStart) + newPXml + xml.substring(pAbsEnd);
    }
    pi++;
  }

  throw new Error('Paragraph not found: p' + paraIdx);
}

function replaceCellText(cellXml: string, value: string, paraIdx: number): string {
  if (paraIdx >= 0) {
    // 특정 문단만 교체
    const pRe = /<P[\s\S]*?<\/P>/g;
    let pm;
    let pi = 0;
    while ((pm = pRe.exec(cellXml))) {
      if (pi === paraIdx) {
        const newP = replaceParaText(pm[0], value);
        return cellXml.substring(0, pm.index) + newP + cellXml.substring(pm.index + pm[0].length);
      }
      pi++;
    }
    throw new Error('Para index out of range: p' + paraIdx);
  }

  // 전체 교체: 모든 P 태그의 CHAR 내용 교체
  // 여러 줄이면 \r\n으로 분리 → 각 P 태그에 배분
  const lines = value.split(/\r?\n/);
  const pRe = /<P[\s\S]*?<\/P>/g;
  const pMatches: RegExpExecArray[] = [];
  let pm;
  while ((pm = pRe.exec(cellXml))) {
    pMatches.push({ ...pm } as unknown as RegExpExecArray);
  }

  if (pMatches.length === 0) return cellXml;

  // 기존 P 태그 수와 줄 수 맞추기
  let result = cellXml;
  if (lines.length <= pMatches.length) {
    // 줄이 적거나 같으면: 앞부터 채우고 나머지 P는 비우기
    for (let i = pMatches.length - 1; i >= 0; i--) {
      const text = i < lines.length ? lines[i] : '';
      const newP = replaceParaText(pMatches[i][0], text);
      result = result.substring(0, pMatches[i].index) + newP + result.substring(pMatches[i].index + pMatches[i][0].length);
    }
  } else {
    // 줄이 더 많으면: 첫 P에 전체를 넣음 (한계)
    const newP = replaceParaText(pMatches[0][0], value.replace(/\r?\n/g, '\r\n'));
    result = result.substring(0, pMatches[0].index) + newP + result.substring(pMatches[0].index + pMatches[0][0].length);
  }

  return result;
}

function replaceParaText(pXml: string, newText: string): string {
  // 첫 번째 TEXT 태그의 속성과 CHAR 속성을 보존
  const firstTextMatch = pXml.match(/<TEXT([^>]*)>([\s\S]*?)<\/TEXT>/);
  const firstEmptyMatch = pXml.match(/<TEXT([^/]*?)\/>/);

  let textAttr = '';
  let charAttr = '';

  if (firstTextMatch) {
    textAttr = firstTextMatch[1];
    // <CHAR .../> (셀프클로징) 과 <CHAR ...>내용</CHAR> 둘 다 처리
    const charOpenMatch = firstTextMatch[2].match(/<CHAR([^/>]*)\/?>/);
    charAttr = charOpenMatch ? charOpenMatch[1] : '';
  } else if (firstEmptyMatch) {
    textAttr = firstEmptyMatch[1];
  }

  // 모든 TEXT 태그 (열린 태그 + 셀프클로징) 제거
  let result = pXml
    .replace(/<TEXT[^>]*>[\s\S]*?<\/TEXT>/g, '')   // <TEXT ...>...</TEXT>
    .replace(/<TEXT[^/]*?\/>/g, '');                 // <TEXT .../>

  // 새 TEXT 태그 하나를 </P> 바로 앞에 삽입
  const newTextTag = '<TEXT' + textAttr + '><CHAR' + charAttr + '>' + escapeXml(newText) + '</CHAR></TEXT>';
  result = result.replace(/<\/P>/, newTextTag + '</P>');

  return result;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
