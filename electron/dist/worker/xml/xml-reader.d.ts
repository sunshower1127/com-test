/**
 * GET 계열 — 문서 구조 읽기
 * 순수 함수: XML 입력 → 문자열 출력 (COM 호출 없음)
 */
import { PageBoundary } from './types';
/** 문서 목차 (크기만, 내용 없음). 드릴다운 + 범위 지정 가능 */
export declare function outline(xml: string, path?: string, pageBoundaries?: PageBoundary[]): string;
/** 간결 구조맵 (범위 지정 가능) */
export declare function get(xml: string, listIdMap: Record<string, number>, path?: string): string;
export declare function styles(xml: string, type?: string, id?: number): string;
/** 기존 structure (= get 전체) */
export declare function structure(xml: string, listIdMap: Record<string, number>): string;
/** 상세 구조맵 — CharShape/ParaShape/BorderFill 포함 */
export declare function structureDetail(xml: string, listIdMap: Record<string, number>): string;
