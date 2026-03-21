/**
 * VM 샌드박스 실행기 + 세이브포인트 자동 롤백
 */
type ComHandle = unknown;
interface ComBridge {
    comGet(handle: ComHandle, prop: string, args?: string[]): unknown;
    comPut(handle: ComHandle, prop: string, value: unknown): void;
    comCall(handle: ComHandle, method: string, args?: string[]): unknown;
    comCallWith(handle: ComHandle, method: string, args: unknown[]): unknown;
}
interface AppHandles {
    excel?: ComHandle;
    hwp?: ComHandle;
    word?: ComHandle;
    ppt?: ComHandle;
}
export interface ExecutionResult {
    success: boolean;
    result?: unknown;
    error?: string;
    stack?: string;
    line?: number;
    code: string;
    logs: string[];
    restored: boolean;
}
export declare function initExecutor(b: ComBridge): void;
/**
 * 세이브포인트 + 실행 + 실패 시 자동 복원
 */
export declare function executeWithSavepoint(code: string, apps: AppHandles): ExecutionResult;
/**
 * 세이브포인트 없이 실행 (새 문서 등)
 */
export declare function executeCode(code: string, apps: AppHandles): ExecutionResult;
export {};
