/**
 * createComProxy — JS Proxy로 COM 객체를 자연스러운 문법으로 접근
 *
 * 사용 예:
 *   const excel = createComProxy(excelHandle);
 *   excel.Visible = true;
 *   excel.Workbooks.Add();
 *   excel.ActiveSheet.Range("A1").Value = "hello";
 */
type ComHandle = unknown;
interface ComBridge {
    comGet(handle: ComHandle, prop: string, args?: string[]): unknown;
    comPut(handle: ComHandle, prop: string, value: unknown): void;
    comCallWith(handle: ComHandle, method: string, args: unknown[]): unknown;
}
export declare function initBridge(b: ComBridge): void;
/**
 * COM 프로퍼티 접근 시 get과 call을 모두 처리해야 함.
 * sheet.Range("A1") — Range는 get이면서 call
 * sheet.Name — Name은 순수 get
 *
 * 해결: get 시 "callable proxy"를 반환.
 * - 함수처럼 호출하면 → comCallWith (메서드/파라미터 프로퍼티)
 * - 프로퍼티 접근하면 → 먼저 comGet 실행 후 결과의 프로퍼티 접근
 * - 값을 읽으면 → comGet 실행 후 원시값 반환
 */
export declare function createComProxy(handle: ComHandle): any;
export {};
