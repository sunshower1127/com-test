/**
 * xml 객체 — 공개 API 파사드
 *
 * 상태 관리 + 각 모듈 위임
 * GET: outline, get(=structure), raw, styles
 * SET: rawSet, setText, insert, insertAfter, insertBefore, addStyle
 */

import { ComBridge, ComHandle } from './types';
import { findElement } from './path-resolver';
import * as reader from './xml-reader';
import * as writer from './xml-writer';
import { mapListIds as doMapListIds } from './list-id';
import { detectPageBoundaries, PageBoundary } from './page-info';

// ──────── 상태 ────────

let _bridge: ComBridge;
let _hwpHandle: ComHandle;
let _xml: string = '';
let _listIdMap: Record<string, number> = {};
let _listIdMapped: boolean = false;
let _pageBoundaries: PageBoundary[] = [];
let _dirty: boolean = false;

// ──────── 초기화 ────────

export function init(bridge: ComBridge): void {
  _bridge = bridge;
}

export function setHandle(handle: ComHandle): void {
  _hwpHandle = handle;
  _xml = '';
  _listIdMap = {};
  _listIdMapped = false;
  _pageBoundaries = [];
  _dirty = false;
}

// ──────── Load / Commit ────────

/** 문서에서 최신 HWPML2X 로드 */
export function load(): string {
  _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
  _dirty = false;
  _listIdMapped = false;  // 문서 변경 가능성 → 재매핑 필요
  return 'XML loaded (' + _xml.length + ' chars)';
}

/** 수정된 XML을 문서에 반영 */
export function commit(): string {
  if (!_dirty || !_xml) return 'No changes to commit.';

  _bridge.comCallWith(_hwpHandle, 'Run', ['SelectAll']);
  _bridge.comCallWith(_hwpHandle, 'Run', ['Delete']);
  _bridge.comCallWith(_hwpHandle, 'SetTextFile', [_xml, 'HWPML2X', '']);

  _dirty = false;
  return 'Committed.';
}

/** xml 블록 시작 시 호출 — 최신 XML 자동 로드 */
export function autoLoad(): void {
  load();
}

/** xml 블록 종료 시 호출 — 변경사항 있으면 자동 커밋 */
export function autoCommit(): void {
  if (_dirty) commit();
}

// ──────── GET 계열 ────────

/** 경량 목차 (크기만, 내용 없음) */
export function outline(path?: string): string {
  if (!_xml) load();
  return reader.outline(_xml, path);
}

/** 간결 구조맵 — 범위 지정 가능 (= 기존 structure 확장) */
export function get(path?: string): string {
  if (!_xml) load();
  // 자동 List ID 매핑 (최초 1회)
  if (!_listIdMapped && _bridge && _hwpHandle) {
    _listIdMap = doMapListIds(_xml, _bridge, _hwpHandle);
    _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
    _listIdMapped = true;
  }
  return reader.get(_xml, _listIdMap, path);
}

/** 스타일 조회 — CharShape/ParaShape */
export function styles(type?: string, id?: number): string {
  if (!_xml) load();
  return reader.styles(_xml, type, id);
}

/** 간결 구조맵 (= get 전체) @deprecated use get() */
export function structure(): string {
  if (!_xml) load();
  return reader.structure(_xml, _listIdMap);
}

/** 상세 구조맵 @deprecated use raw() */
export function structureDetail(): string {
  if (!_xml) load();
  return reader.structureDetail(_xml, _listIdMap);
}

/** 특정 경로의 원본 XML (화이트리스트 정제) */
export function raw(path: string): string {
  if (!_xml) load();
  const loc = findElement(_xml, path);
  return loc.xml;
}

// ──────── SET 계열 ────────

/** 원본 XML 교체 (빈 문자열 = 삭제) */
export function rawSet(path: string, newXml: string): string {
  if (!_xml) load();

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
export function setText(path: string, text: string): string {
  if (!_xml) load();
  _xml = writer.setText(_xml, path, text);
  _dirty = true;
  return 'setText ' + path + ' applied.';
}

/** 지정 경로 뒤에 삽입 */
export function insert(path: string, xmlString: string): string {
  if (!_xml) load();
  _xml = writer.insertAfter(_xml, path, xmlString);
  _dirty = true;
  return 'insert after ' + path + ' applied.';
}

/** insert의 별칭 */
export function insertAfter(path: string, xmlString: string): string {
  return insert(path, xmlString);
}

/** 지정 경로 앞에 삽입 */
export function insertBefore(path: string, xmlString: string): string {
  if (!_xml) load();
  _xml = writer.insertBefore(_xml, path, xmlString);
  _dirty = true;
  return 'insertBefore ' + path + ' applied.';
}

/** HEAD에 새 CharShape/ParaShape 추가 */
export function addStyle(type: string, props: Record<string, unknown>): number {
  if (!_xml) load();
  const result = writer.addStyle(_xml, type, props);
  _xml = result.xml;
  _dirty = true;
  return result.id;
}

// ──────── List ID ────────

/** 모든 표 셀의 List ID 매핑 (COM 필요) */
export function mapListIds(): string {
  if (!_xml) load();
  _listIdMap = doMapListIds(_xml, _bridge, _hwpHandle);
  // 매핑 후 최신 XML 재로드
  _xml = String(_bridge.comCallWith(_hwpHandle, 'GetTextFile', ['HWPML2X', '']));
  const mapped = Object.keys(_listIdMap).length;
  return 'Mapped ' + mapped + ' cells.';
}

// ──────── 하위 호환 (deprecated) ────────

/** @deprecated — use setText instead */
export function set(path: string, value: string): string {
  return setText(path, value);
}

/** @deprecated — use autoLoad/autoCommit */
export function getXml(): string {
  return load();
}

/** @deprecated — use append via raw */
export function append(path: string, text: string): string {
  // 기존 호환: 큐잉 방식 대신 즉시 적용
  if (!_xml) load();
  const loc = findElement(_xml, path);
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
