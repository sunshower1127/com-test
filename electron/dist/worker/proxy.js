"use strict";
/**
 * createComProxy — JS Proxy로 COM 객체를 자연스러운 문법으로 접근
 *
 * 사용 예:
 *   const excel = createComProxy(excelHandle);
 *   excel.Visible = true;
 *   excel.Workbooks.Add();
 *   excel.ActiveSheet.Range("A1").Value = "hello";
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.initBridge = initBridge;
exports.createComProxy = createComProxy;
let bridge;
function initBridge(b) {
    bridge = b;
}
function isComHandle(val) {
    // napi External objects have a specific internal type
    // In practice, they are objects that aren't plain JS objects
    return val !== null && val !== undefined && typeof val === 'object' && !Array.isArray(val);
}
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
function createComProxy(handle) {
    return new Proxy(() => { }, {
        get(_target, prop) {
            if (typeof prop === 'symbol')
                return undefined;
            if (prop === '_handle')
                return handle;
            if (prop === 'then')
                return undefined; // Promise 체이닝 방지
            if (prop === 'toJSON')
                return () => '[COM Object]';
            if (prop === 'valueOf') {
                // 원시값 변환 시도 시 comGet 실행
                return () => bridge.comGet(handle, 'Value');
            }
            // "callable property" 반환 — 호출하면 call, 접근하면 get
            return createCallableProperty(handle, prop);
        },
        set(_target, prop, value) {
            if (typeof prop === 'symbol')
                return false;
            // Proxy 객체가 넘어오면 handle 추출
            const rawValue = value?._handle ?? value;
            bridge.comPut(handle, prop, rawValue);
            return true;
        },
        apply(_target, _thisArg, args) {
            // handle 자체가 호출된 경우 (부모의 method call)
            // 이 경우는 createCallableProperty에서 처리
            return undefined;
        },
    });
}
function createCallableProperty(parentHandle, prop) {
    // 이 함수가 반환하는 Proxy는:
    // 1. 함수로 호출 가능 → comCallWith(parent, prop, args)
    // 2. 프로퍼티 접근 가능 → comGet(parent, prop) 후 결과의 프로퍼티
    // 3. 값 할당 가능 → 부모의 sub-property put
    // comGet 결과를 캐싱하여 동일 handle 재사용
    // (Properties 패턴: get → modify → set back에서 같은 객체를 유지)
    let _cached = null;
    function resolve() {
        if (!_cached)
            _cached = bridge.comGet(parentHandle, prop);
        return _cached;
    }
    return new Proxy(() => { }, {
        apply(_target, _thisArg, args) {
            // sheet.Range("A1") → comCallWith(sheet, "Range", ["A1"])
            const rawArgs = args.map(a => a?._handle ?? a);
            const result = bridge.comCallWith(parentHandle, prop, rawArgs);
            return wrapResult(result);
        },
        get(_target, subProp) {
            // Symbol.toPrimitive — 문자열 연결, 산술 연산 시 자동 호출
            if (subProp === Symbol.toPrimitive) {
                return () => {
                    const result = resolve();
                    if (isComHandle(result))
                        return '[COM Object]';
                    return result;
                };
            }
            if (typeof subProp === 'symbol')
                return undefined;
            if (subProp === '_handle') {
                return resolve();
            }
            if (subProp === 'then')
                return undefined;
            if (subProp === 'valueOf' || subProp === 'toString') {
                return () => {
                    const result = resolve();
                    if (isComHandle(result))
                        return '[COM Object]';
                    return result;
                };
            }
            // 캐싱된 comGet 결과 사용
            const result = resolve();
            if (isComHandle(result)) {
                const proxy = createComProxy(result);
                return proxy[subProp];
            }
            // 원시값 → JS 네이티브 프로퍼티 접근
            return result?.[subProp];
        },
        set(_target, subProp, value) {
            if (typeof subProp === 'symbol')
                return false;
            const result = resolve();
            if (isComHandle(result)) {
                const rawValue = value?._handle ?? value;
                bridge.comPut(result, subProp, rawValue);
                return true;
            }
            return false;
        },
    });
}
function wrapResult(result) {
    if (isComHandle(result)) {
        return createComProxy(result);
    }
    return result;
}
