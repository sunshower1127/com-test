"use strict";
/**
 * VM 샌드박스 실행기 + 세이브포인트 자동 롤백
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
exports.initExecutor = initExecutor;
exports.executeWithSavepoint = executeWithSavepoint;
exports.executeCode = executeCode;
const vm = __importStar(require("vm"));
const fs = __importStar(require("fs"));
const os = __importStar(require("os"));
const path = __importStar(require("path"));
const proxy_1 = require("./proxy");
let bridge;
function initExecutor(b) {
    bridge = b;
}
/**
 * LLM 코드를 VM 샌드박스에서 실행
 */
function runInSandbox(code, apps) {
    const logs = [];
    const sandbox = {
        console: {
            log: (...args) => logs.push(args.map(String).join(' ')),
            error: (...args) => logs.push('[ERROR] ' + args.map(String).join(' ')),
        },
        result: null,
    };
    // COM proxy 객체를 sandbox에 주입
    if (apps.excel) {
        sandbox.excel = (0, proxy_1.createComProxy)(apps.excel);
    }
    if (apps.hwp) {
        sandbox.hwp = (0, proxy_1.createComProxy)(apps.hwp);
    }
    try {
        vm.runInNewContext(code, sandbox, {
            timeout: 30000, // 30초
            displayErrors: true,
        });
        return {
            success: true,
            result: sandbox.result,
            code,
            logs,
            restored: false,
        };
    }
    catch (e) {
        const err = e;
        const lineMatch = err.stack?.match(/:(\d+)/);
        return {
            success: false,
            error: err.message,
            stack: err.stack,
            line: lineMatch ? parseInt(lineMatch[1], 10) : undefined,
            code,
            logs,
            restored: false,
        };
    }
}
/**
 * 세이브포인트 + 실행 + 실패 시 자동 복원
 */
function executeWithSavepoint(code, apps) {
    const tempDir = os.tmpdir();
    const timestamp = Date.now();
    let savepointPath = null;
    let activeApp = null;
    // 1. 세이브포인트 생성
    try {
        if (apps.excel) {
            savepointPath = path.join(tempDir, `savepoint_excel_${timestamp}.xlsx`);
            bridge.comCallWith(apps.excel, 'SaveCopyAs', [savepointPath]);
            activeApp = 'excel';
        }
        else if (apps.hwp) {
            savepointPath = path.join(tempDir, `savepoint_hwp_${timestamp}.hwp`);
            bridge.comCallWith(apps.hwp, 'SaveAs', [savepointPath, 'HWP', '']);
            activeApp = 'hwp';
        }
    }
    catch {
        // 세이브포인트 생성 실패 — 새 문서일 수 있음. 그냥 진행.
        savepointPath = null;
    }
    // 2. 코드 실행
    const result = runInSandbox(code, apps);
    // 3. 실패 시 복원
    if (!result.success && savepointPath && activeApp) {
        try {
            if (activeApp === 'excel' && apps.excel) {
                bridge.comPut(apps.excel, 'DisplayAlerts', false);
                bridge.comCallWith(apps.excel, 'Close', [false]);
                // Workbooks.Open으로 복원
                const wbs = bridge.comGet(apps.excel, 'Workbooks');
                bridge.comCallWith(wbs, 'Open', [savepointPath]);
                bridge.comPut(apps.excel, 'DisplayAlerts', true);
            }
            else if (activeApp === 'hwp' && apps.hwp) {
                bridge.comCallWith(apps.hwp, 'Clear', [1]);
                bridge.comCallWith(apps.hwp, 'Open', [savepointPath]);
            }
            result.restored = true;
        }
        catch {
            // 복원 실패 — 에러에 기록
            result.error = (result.error || '') + ' [WARN: savepoint restore failed]';
        }
        // 임시파일 정리
        try {
            fs.unlinkSync(savepointPath);
        }
        catch { /* ignore */ }
    }
    else if (savepointPath) {
        // 성공 시 임시파일 정리
        try {
            fs.unlinkSync(savepointPath);
        }
        catch { /* ignore */ }
    }
    return result;
}
/**
 * 세이브포인트 없이 실행 (새 문서 등)
 */
function executeCode(code, apps) {
    return runInSandbox(code, apps);
}
