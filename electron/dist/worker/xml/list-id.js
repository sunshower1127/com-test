"use strict";
/**
 * List ID 매핑 — COM 호출로 각 표 셀의 런타임 List ID를 알아내기
 * 현재: 모든 셀에 마커 삽입 → RepeatFind → 제거
 * TODO: 표당 첫 셀만 마커, 나머지는 +1 (간소화)
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.mapListIds = mapListIds;
const path_resolver_1 = require("./path-resolver");
/** 모든 표 셀의 List ID를 매핑 */
function mapListIds(xml, bridge, hwpHandle) {
    const tables = (0, path_resolver_1.parseTables)(xml);
    const markerPrefix = '{{#LID_';
    const paths = [];
    const listIdMap = {};
    // 1. 각 셀에 유니크 마커 append
    let modifiedXml = xml;
    for (let ti = 0; ti < tables.length; ti++) {
        tables[ti].forEach((row, ri) => {
            row.forEach((cell) => {
                const path = 't' + ti + '.r' + ri + '.c' + cell.col;
                const marker = markerPrefix + path + '}}';
                paths.push(path);
                modifiedXml = appendToCellXml(modifiedXml, ['t' + ti, 'r' + ri, 'c' + cell.col], marker);
            });
        });
    }
    // 2. 문서에 반영
    bridge.comCallWith(hwpHandle, 'Run', ['SelectAll']);
    bridge.comCallWith(hwpHandle, 'Run', ['Delete']);
    bridge.comCallWith(hwpHandle, 'SetTextFile', [modifiedXml, 'HWPML2X', '']);
    // 3. 각 마커를 RepeatFind로 찾아서 List ID 기록
    bridge.comCallWith(hwpHandle, 'Run', ['MoveDocBegin']);
    for (const path of paths) {
        const marker = markerPrefix + path + '}}';
        const hAction = bridge.comGet(hwpHandle, 'HAction');
        const hParamSet = bridge.comGet(hwpHandle, 'HParameterSet');
        const hFindReplace = bridge.comGet(hParamSet, 'HFindReplace');
        const hSet = bridge.comGet(hFindReplace, 'HSet');
        bridge.comCallWith(hAction, 'GetDefault', ['RepeatFind', hSet]);
        bridge.comPut(hFindReplace, 'FindString', marker);
        bridge.comPut(hFindReplace, 'IgnoreMessage', 1);
        const found = bridge.comCallWith(hAction, 'Execute', ['RepeatFind', hSet]);
        if (found) {
            const ps = bridge.comCallWith(hwpHandle, 'GetPosBySet', []);
            const listId = bridge.comCallWith(ps, 'Item', ['List']);
            if (typeof listId === 'number') {
                listIdMap[path] = listId;
            }
        }
    }
    // 4. 마커 전체 제거 — HWPML2X에서 직접 제거 후 재로드
    let cleanXml = String(bridge.comCallWith(hwpHandle, 'GetTextFile', ['HWPML2X', '']));
    cleanXml = cleanXml.replace(/\{\{#LID_[^}]*\}\}/g, '');
    bridge.comCallWith(hwpHandle, 'Run', ['SelectAll']);
    bridge.comCallWith(hwpHandle, 'Run', ['Delete']);
    bridge.comCallWith(hwpHandle, 'SetTextFile', [cleanXml, 'HWPML2X', '']);
    return listIdMap;
}
/** 셀 XML의 마지막 CHAR 앞에 텍스트 추가 */
function appendToCellXml(xml, parts, text) {
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
                        const ca = (0, path_resolver_1.getAttr)(cm[0], 'ColAddr');
                        if (ca === colIdx) {
                            const cellXml = cm[0];
                            let newCell;
                            const lastCharClose = cellXml.lastIndexOf('</CHAR>');
                            if (lastCharClose >= 0) {
                                newCell = cellXml.substring(0, lastCharClose) + (0, path_resolver_1.escapeXml)(text) + cellXml.substring(lastCharClose);
                            }
                            else {
                                const lastTextClose = cellXml.lastIndexOf('</TEXT>');
                                if (lastTextClose >= 0) {
                                    newCell = cellXml.substring(0, lastTextClose) + '<CHAR>' + (0, path_resolver_1.escapeXml)(text) + '</CHAR>' + cellXml.substring(lastTextClose);
                                }
                                else {
                                    const selfTextRe = /<TEXT([^/]*?)\/>/;
                                    const stm = selfTextRe.exec(cellXml);
                                    if (stm) {
                                        const replacement = '<TEXT' + stm[1] + '><CHAR>' + (0, path_resolver_1.escapeXml)(text) + '</CHAR></TEXT>';
                                        newCell = cellXml.substring(0, stm.index) + replacement + cellXml.substring(stm.index + stm[0].length);
                                    }
                                    else {
                                        newCell = cellXml;
                                    }
                                }
                            }
                            const absStart = tableStart + rm.index + cm.index;
                            const absEnd = absStart + cellXml.length;
                            return xml.substring(0, absStart) + newCell + xml.substring(absEnd);
                        }
                    }
                    return xml;
                }
                ri++;
            }
            return xml;
        }
        ti++;
    }
    return xml;
}
