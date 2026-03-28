"use strict";
/**
 * xml 객체 — 공개 API 파사드
 *
 * 상태 관리 + 각 모듈 위임
 * GET: outline, get(=structure), raw, styles
 * SET: rawSet, setText, insert, insertAfter, insertBefore, addStyle
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.init = init;
exports.setHandle = setHandle;
exports.load = load;
exports.commit = commit;
exports.autoLoad = autoLoad;
exports.autoCommit = autoCommit;
exports.outline = outline;
exports.get = get;
exports.styles = styles;
exports.structure = structure;
exports.structureDetail = structureDetail;
exports.raw = raw;
exports.rawSet = rawSet;
exports.setText = setText;
exports.insert = insert;
exports.insertAfter = insertAfter;
exports.insertBefore = insertBefore;
exports.addStyle = addStyle;
exports.mapListIds = mapListIds;
exports.set = set;
exports.getXml = getXml;
exports.append = append;
const path_resolver_1 = require("./path-resolver");
const reader = __importStar(require("./xml-reader"));
const writer = __importStar(require("./xml-writer"));
const list_id_1 = require("./list-id");
const page_info_1 = require("./page-info");
// ──────── 상태 ────────
let _bridge;
let _hwpHandle;
let _xml = '';
let _listIdMap = {};
let _mappedTables = new Set();
let _pageBoundaries = [];
let _dirty = false;
// ──────── 초기화 ────────
function init(bridge) {
    _bridge = bridge;
}
function setHandle(handle) {
    _hwpHandle = handle;
    _xml = '';
    _listIdMap = {};
    _mappedTables = new Set();
    _pageBoundaries = [];
    _dirty = false;
}
// ──────── Load / Commit ────────
/** 문서에서 최신 HWPML2X 로드 */
function load() {
    _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
    _dirty = false;
    _mappedTables = new Set(); // 문서 변경 가능성 → 재매핑 필요
    return 'XML loaded (' + _xml.length + ' chars)';
}
/** 수정된 XML을 문서에 반영 */
function commit() {
    if (!_dirty || !_xml)
        return 'No changes to commit.';
    _bridge.comCallWith(_hwpHandle, 'Run', ['SelectAll']);
    _bridge.comCallWith(_hwpHandle, 'Run', ['Delete']);
    _bridge.comCallWith(_hwpHandle, 'SetTextFile', [_xml, 'HWPML2X', '']);
    _dirty = false;
    return 'Committed.';
}
/** xml 블록 시작 시 호출 — 최신 XML 자동 로드 */
function autoLoad() {
    load();
}
/** xml 블록 종료 시 호출 — 변경사항 있으면 자동 커밋 */
function autoCommit() {
    if (_dirty)
        commit();
}
// ──────── GET 계열 ────────
/** 경량 목차 (크기만, 내용 없음, 페이지 정보 포함) */
function outline(path) {
    if (!_xml)
        load();
    // 페이지 경계 감지 (최초 1회)
    if (_pageBoundaries.length === 0 && _bridge && _hwpHandle) {
        _pageBoundaries = (0, page_info_1.detectPageBoundaries)(_bridge, _hwpHandle);
    }
    return reader.outline(_xml, path, _pageBoundaries);
}
/** 간결 구조맵 — 범위 지정 가능 (= 기존 structure 확장) */
function get(path) {
    if (!_xml)
        load();
    // path에서 필요한 테이블만 lazy 매핑
    if (_bridge && _hwpHandle) {
        const needed = extractTableIndices(path);
        const unmapped = needed.filter(ti => !_mappedTables.has(ti));
        if (unmapped.length > 0) {
            const newMap = (0, list_id_1.mapListIds)(_xml, _bridge, _hwpHandle, unmapped);
            Object.assign(_listIdMap, newMap);
            _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
            for (const ti of unmapped)
                _mappedTables.add(ti);
        }
    }
    return reader.get(_xml, _listIdMap, path);
}
/** path에서 필요한 테이블 인덱스 추출 */
function extractTableIndices(path) {
    if (!path) {
        // 전체 조회 — 모든 테이블 필요
        const tableRe = /<TABLE/g;
        let count = 0;
        while (tableRe.exec(_xml))
            count++;
        return Array.from({ length: count }, (_, i) => i);
    }
    // t0, t0.r1, t0.r1.c3, t0~t2, t0.r0~t1.r3 등에서 테이블 인덱스 추출
    const indices = new Set();
    const tRe = /\bt(\d+)/g;
    let m;
    while ((m = tRe.exec(path))) {
        indices.add(parseInt(m[1]));
    }
    // 범위: t0~t3 → 0,1,2,3
    const rangeRe = /t(\d+)[^~]*~[^t]*t(\d+)/;
    const rm = rangeRe.exec(path);
    if (rm) {
        const from = parseInt(rm[1]);
        const to = parseInt(rm[2]);
        for (let i = from; i <= to; i++)
            indices.add(i);
    }
    return Array.from(indices);
}
/** 스타일 조회 — CharShape/ParaShape */
function styles(type, id) {
    if (!_xml)
        load();
    return reader.styles(_xml, type, id);
}
/** 간결 구조맵 (= get 전체) @deprecated use get() */
function structure() {
    if (!_xml)
        load();
    return reader.structure(_xml, _listIdMap);
}
/** 상세 구조맵 @deprecated use raw() */
function structureDetail() {
    if (!_xml)
        load();
    return reader.structureDetail(_xml, _listIdMap);
}
/** 특정 경로의 원본 XML (화이트리스트 정제) */
function raw(path) {
    if (!_xml)
        load();
    const loc = (0, path_resolver_1.findElement)(_xml, path);
    return loc.xml;
}
// ──────── SET 계열 ────────
/** 원본 XML 교체 (빈 문자열 = 삭제) */
function rawSet(path, newXml) {
    if (!_xml)
        load();
    // 간단한 XML 태그 균형 검증
    if (newXml) {
        const openTags = (newXml.match(/<[a-zA-Z][^/>]*>/g) || []).length;
        const closeTags = (newXml.match(/<\/[a-zA-Z][^>]*>/g) || []).length;
        const selfClose = (newXml.match(/<[a-zA-Z][^>]*\/>/g) || []).length;
        if (openTags !== closeTags + selfClose && openTags !== closeTags) {
            console.warn('XML 태그 불균형 경고: open=' + openTags + ' close=' + closeTags + ' selfClose=' + selfClose);
        }
    }
    _xml = writer.rawSet(_xml, path, newXml);
    _dirty = true;
    return 'SET ' + path + ' applied.';
}
/** 텍스트만 교체 (서식 유지) */
function setText(path, text) {
    if (!_xml)
        load();
    _xml = writer.setText(_xml, path, text);
    _dirty = true;
    return 'setText ' + path + ' applied.';
}
/** 지정 경로 뒤에 삽입 */
function insert(path, xmlString) {
    if (!_xml)
        load();
    _xml = writer.insertAfter(_xml, path, xmlString);
    _dirty = true;
    return 'insert after ' + path + ' applied.';
}
/** insert의 별칭 */
function insertAfter(path, xmlString) {
    return insert(path, xmlString);
}
/** 지정 경로 앞에 삽입 */
function insertBefore(path, xmlString) {
    if (!_xml)
        load();
    _xml = writer.insertBefore(_xml, path, xmlString);
    _dirty = true;
    return 'insertBefore ' + path + ' applied.';
}
/** HEAD에 새 CharShape/ParaShape 추가 */
function addStyle(type, props) {
    if (!_xml)
        load();
    const result = writer.addStyle(_xml, type, props);
    _xml = result.xml;
    _dirty = true;
    return result.id;
}
// ──────── List ID ────────
/** 모든 표 셀의 List ID 매핑 (COM 필요) */
function mapListIds() {
    if (!_xml)
        load();
    _listIdMap = (0, list_id_1.mapListIds)(_xml, _bridge, _hwpHandle);
    _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
    // 전체 매핑됨
    const tableRe = /<TABLE/g;
    let count = 0;
    while (tableRe.exec(_xml))
        count++;
    _mappedTables = new Set(Array.from({ length: count }, (_, i) => i));
    const mapped = Object.keys(_listIdMap).length;
    return 'Mapped ' + mapped + ' cells.';
}
// ──────── 하위 호환 (deprecated) ────────
/** @deprecated — use setText instead */
function set(path, value) {
    return setText(path, value);
}
/** @deprecated — use autoLoad/autoCommit */
function getXml() {
    return load();
}
/** @deprecated — use append via raw */
function append(path, text) {
    // 기존 호환: 큐잉 방식 대신 즉시 적용
    if (!_xml)
        load();
    const loc = (0, path_resolver_1.findElement)(_xml, path);
    const cellXml = loc.xml;
    const lastCharClose = cellXml.lastIndexOf('</CHAR>');
    if (lastCharClose >= 0) {
        const newCell = cellXml.substring(0, lastCharClose) +
            text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') +
            cellXml.substring(lastCharClose);
        _xml = _xml.substring(0, loc.start) + newCell + _xml.substring(loc.end);
        _dirty = true;
    }
    return 'appended to ' + path;
}
