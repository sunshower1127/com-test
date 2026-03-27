/**
 * xml 객체 — 공개 API 파사드
 *
 * 상태 관리 + 각 모듈 위임
 * GET: outline, get(=structure), raw, styles
 * SET: rawSet, setText, insert, insertAfter, insertBefore, addStyle
 */
import { ComBridge, ComHandle } from './types';
export declare function init(bridge: ComBridge): void;
export declare function setHandle(handle: ComHandle): void;
/** 문서에서 최신 HWPML2X 로드 */
export declare function load(): string;
/** 수정된 XML을 문서에 반영 */
export declare function commit(): string;
/** xml 블록 시작 시 호출 — 최신 XML 자동 로드 */
export declare function autoLoad(): void;
/** xml 블록 종료 시 호출 — 변경사항 있으면 자동 커밋 */
export declare function autoCommit(): void;
/** 간결 구조맵 (= 기존 structure) */
export declare function structure(): string;
/** 상세 구조맵 (= 기존 structureDetail) */
export declare function structureDetail(): string;
/** 특정 경로의 원본 XML (화이트리스트 정제) */
export declare function raw(path: string): string;
/** 원본 XML 교체 (빈 문자열 = 삭제) */
export declare function rawSet(path: string, newXml: string): string;
/** 텍스트만 교체 (서식 유지) */
export declare function setText(path: string, text: string): string;
/** 지정 경로 뒤에 삽입 */
export declare function insert(path: string, xmlString: string): string;
/** insert의 별칭 */
export declare function insertAfter(path: string, xmlString: string): string;
/** 지정 경로 앞에 삽입 */
export declare function insertBefore(path: string, xmlString: string): string;
/** 모든 표 셀의 List ID 매핑 (COM 필요) */
export declare function mapListIds(): string;
/** @deprecated — use setText instead */
export declare function set(path: string, value: string): string;
/** @deprecated — use autoLoad/autoCommit */
export declare function getXml(): string;
/** @deprecated — use append via raw */
export declare function append(path: string, text: string): string;
