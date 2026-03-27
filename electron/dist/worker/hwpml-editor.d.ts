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
export declare function initHwpmlEditor(bridge: ComBridge): void;
export declare function setHwpHandle(handle: ComHandle): void;
/** 현재 문서에서 HWPML2X 로드 */
export declare function getXml(): string;
/** 변경 등록 */
export declare function set(path: string, value: string): string;
/** 등록된 변경사항을 XML에 적용하고 문서에 반영 */
export declare function commit(): string;
/** 셀 텍스트 끝에 텍스트 추가 (원본 보존) */
export declare function append(path: string, text: string): string;
/** 모든 표 셀에 마커를 append → RepeatFind로 List ID 매핑 → 마커 제거 */
export declare function mapListIds(): string;
/** 특정 경로의 원본 XML 반환 */
export declare function raw(path: string): string;
/** 특정 경로의 XML을 교체하고 문서에 반영 */
export declare function rawSet(path: string, newXml: string): string;
/** 간결 구조맵 — 스타일 제거, 텍스트와 구조만 */
export declare function structure(): string;
/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 정제 XML */
export declare function structureDetail(): string;
export {};
