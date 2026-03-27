"use strict";
/**
 * SET 계열 — XML 수정
 * 순수 함수: XML 입력 → 수정된 XML 출력 (COM 호출 없음)
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.rawSet = rawSet;
exports.setText = setText;
exports.insertAfter = insertAfter;
exports.insertBefore = insertBefore;
exports.addStyle = addStyle;
const path_resolver_1 = require("./path-resolver");
/** 특정 경로의 XML을 교체. 빈 문자열이면 삭제 */
function rawSet(xml, path, newXml) {
    const loc = (0, path_resolver_1.findElement)(xml, path);
    return xml.substring(0, loc.start) + newXml + xml.substring(loc.end);
}
/** 텍스트만 교체 (기존 서식 구조 유지) */
function setText(xml, path, value) {
    return applyChange(xml, path, value);
}
/** 지정 경로 뒤에 새 노드 삽입 */
function insertAfter(xml, path, newXml) {
    const loc = (0, path_resolver_1.findElement)(xml, path);
    return xml.substring(0, loc.end) + newXml + xml.substring(loc.end);
}
/** 지정 경로 앞에 새 노드 삽입 */
function insertBefore(xml, path, newXml) {
    const loc = (0, path_resolver_1.findElement)(xml, path);
    return xml.substring(0, loc.start) + newXml + xml.substring(loc.start);
}
/** HEAD 섹션에 새 CharShape/ParaShape 추가, 할당된 인덱스 반환 */
function addStyle(xml, type, props) {
    // 기존 스타일 수 세기
    const re = new RegExp('<' + type + '\\b[^>]*/?>', 'g');
    let count = 0;
    let m;
    while ((m = re.exec(xml)))
        count++;
    // 속성 문자열 생성
    const attrStr = Object.entries(props)
        .map(([k, v]) => k + '="' + String(v) + '"')
        .join(' ');
    const newTag = '<' + type + ' ' + attrStr + '/>';
    // 마지막 해당 태그 뒤에 삽입
    const lastRe = new RegExp('<' + type + '\\b[^>]*/?>(?!.*<' + type + '\\b)', 's');
    const lastMatch = lastRe.exec(xml);
    if (lastMatch) {
        const insertPos = lastMatch.index + lastMatch[0].length;
        return {
            xml: xml.substring(0, insertPos) + newTag + xml.substring(insertPos),
            id: count,
        };
    }
    // 해당 태그가 없으면 HEAD 닫기 태그 앞에 삽입
    const headClose = xml.indexOf('</HEAD>');
    if (headClose >= 0) {
        return {
            xml: xml.substring(0, headClose) + newTag + xml.substring(headClose),
            id: 0,
        };
    }
    throw new Error('HEAD section not found');
}
// ──────── 내부: 텍스트 교체 로직 ────────
function applyChange(xml, path, value) {
    const parts = path.split('.');
    if (parts[0].startsWith('t')) {
        return applyTableChange(xml, parts, value);
    }
    else if (parts[0].startsWith('p')) {
        return applyParaChange(xml, parts, value);
    }
    throw new Error('Unknown path: ' + path);
}
function applyTableChange(xml, parts, value) {
    const tableIdx = parseInt(parts[0].substring(1));
    const rowIdx = parseInt(parts[1].substring(1));
    const colIdx = parseInt(parts[2].substring(1));
    const paraIdx = parts.length > 3 ? parseInt(parts[3].substring(1)) : -1;
    const tableRe = /<TABLE[^>]*>[\s\S]*?<\/TABLE>/g;
    let tm;
    let ti = 0;
    while ((tm = tableRe.exec(xml))) {
        if (ti === tableIdx) {
            const tableStart = tm.index;
            const tableXml = tm[0];
            const rowRe = /<ROW[^>]*>[\s\S]*?<\/ROW>/g;
            let rm;
            let ri = 0;
            while ((rm = rowRe.exec(tableXml))) {
                if (ri === rowIdx) {
                    const cellRe = /<CELL\b[^>]*?>[\s\S]*?<\/CELL>/g;
                    let cm = null;
                    let cellMatch = null;
                    while ((cellMatch = cellRe.exec(rm[0]))) {
                        const ca = (0, path_resolver_1.getAttr)(cellMatch[0], 'ColAddr');
                        if (ca === colIdx) {
                            cm = cellMatch;
                            break;
                        }
                    }
                    if (!cm)
                        throw new Error('Cell not found: t' + tableIdx + '.r' + rowIdx + '.c' + colIdx);
                    const cellXml = cm[0];
                    const newCellXml = replaceCellText(cellXml, value, paraIdx);
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
function applyParaChange(xml, parts, value) {
    const paraIdx = parseInt(parts[0].substring(1));
    const loc = (0, path_resolver_1.findParagraph)(xml, paraIdx);
    const newPXml = replaceParaText(loc.xml, value);
    return xml.substring(0, loc.start) + newPXml + xml.substring(loc.end);
}
function replaceCellText(cellXml, value, paraIdx) {
    if (paraIdx >= 0) {
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
    const lines = value.split(/\r?\n/);
    const pRe = /<P[\s\S]*?<\/P>/g;
    const pMatches = [];
    let pm;
    while ((pm = pRe.exec(cellXml))) {
        pMatches.push({ ...pm });
    }
    if (pMatches.length === 0)
        return cellXml;
    let result = cellXml;
    if (lines.length <= pMatches.length) {
        for (let i = pMatches.length - 1; i >= 0; i--) {
            const text = i < lines.length ? lines[i] : '';
            const newP = replaceParaText(pMatches[i][0], text);
            result = result.substring(0, pMatches[i].index) + newP + result.substring(pMatches[i].index + pMatches[i][0].length);
        }
    }
    else {
        const newP = replaceParaText(pMatches[0][0], value.replace(/\r?\n/g, '\r\n'));
        result = result.substring(0, pMatches[0].index) + newP + result.substring(pMatches[0].index + pMatches[0][0].length);
    }
    return result;
}
function replaceParaText(pXml, newText) {
    const firstTextMatch = pXml.match(/<TEXT([^>]*)>([\s\S]*?)<\/TEXT>/);
    const firstEmptyMatch = pXml.match(/<TEXT([^/]*?)\/>/);
    let textAttr = '';
    let charAttr = '';
    if (firstTextMatch) {
        textAttr = firstTextMatch[1];
        const charOpenMatch = firstTextMatch[2].match(/<CHAR([^/>]*)\/?>/);
        charAttr = charOpenMatch ? charOpenMatch[1] : '';
    }
    else if (firstEmptyMatch) {
        textAttr = firstEmptyMatch[1];
    }
    let result = pXml
        .replace(/<TEXT[^>]*>[\s\S]*?<\/TEXT>/g, '')
        .replace(/<TEXT[^/]*?\/>/g, '');
    const newTextTag = '<TEXT' + textAttr + '><CHAR' + charAttr + '>' + (0, path_resolver_1.escapeXml)(newText) + '</CHAR></TEXT>';
    result = result.replace(/<\/P>/, newTextTag + '</P>');
    return result;
}
