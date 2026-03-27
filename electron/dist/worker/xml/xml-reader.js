"use strict";
/**
 * GET 계열 — 문서 구조 읽기
 * 순수 함수: XML 입력 → 문자열 출력 (COM 호출 없음)
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.structure = structure;
exports.structureDetail = structureDetail;
const path_resolver_1 = require("./path-resolver");
const xml_cleaner_1 = require("./xml-cleaner");
/** 간결 구조맵 — 스타일 제거, 텍스트+구조만 */
function structure(xml, listIdMap) {
    const { section } = (0, path_resolver_1.extractSection)(xml);
    const elements = (0, path_resolver_1.collectDocElements)(section);
    const out = [];
    let pIdx = 0;
    let tIdx = 0;
    for (const el of elements) {
        if (el.type === 'p') {
            out.push((0, xml_cleaner_1.compactParagraph)(el.content, pIdx));
            pIdx++;
        }
        else {
            out.push((0, xml_cleaner_1.compactTable)(el.content, tIdx, listIdMap));
            tIdx++;
        }
    }
    return out.join('\n');
}
/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 */
function structureDetail(xml, listIdMap) {
    const { section } = (0, path_resolver_1.extractSection)(xml);
    const elements = (0, path_resolver_1.collectDocElements)(section);
    const out = [];
    let pIdx = 0;
    let tIdx = 0;
    for (const el of elements) {
        if (el.type === 'p') {
            out.push((0, xml_cleaner_1.cleanParagraph)(el.content, 'p' + pIdx, 0));
            pIdx++;
        }
        else {
            out.push((0, xml_cleaner_1.cleanTable)(el.content, tIdx, 0, listIdMap));
            tIdx++;
        }
    }
    return out.join('\n');
}
