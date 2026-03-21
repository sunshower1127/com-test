"use strict";
/**
 * Worker process entry point
 *
 * 실행 방법:
 * - Electron utilityProcess: main에서 fork
 * - 단독 테스트: node dist/worker/index.js
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
const path = __importStar(require("path"));
const proxy_1 = require("./proxy");
const executor_1 = require("./executor");
// .node 네이티브 모듈 로드
const nativePath = path.resolve(__dirname, '../../native/com_bridge_node.node');
let bridge;
try {
    bridge = require(nativePath);
}
catch {
    // fallback: 빌드 output에서 로드
    const fallbackPath = path.resolve(__dirname, '../../../target/debug/com_bridge_node.node');
    bridge = require(fallbackPath);
}
// COM 초기화
bridge.comInit();
// Proxy & Executor에 bridge 주입
(0, proxy_1.initBridge)(bridge);
(0, executor_1.initExecutor)(bridge);
// 앱 핸들 저장소
const apps = {};
function handleMessage(msg) {
    switch (msg.type) {
        case 'launch': {
            const progId = msg.app === 'excel' ? 'Excel.Application' : 'HWPFrame.HwpObject';
            const handle = bridge.comCreate(progId);
            apps[msg.app] = handle;
            // visible 설정
            if (msg.app === 'excel') {
                bridge.comPut(handle, 'Visible', true);
            }
            else {
                const wins = bridge.comGet(handle, 'XHwpWindows');
                const win0 = bridge.comCallWith(wins, 'Item', [0]);
                bridge.comPut(win0, 'Visible', true);
            }
            return {
                type: 'status',
                excel: !!apps.excel,
                hwp: !!apps.hwp,
            };
        }
        case 'quit': {
            const handle = apps[msg.app];
            if (handle) {
                if (msg.app === 'excel') {
                    bridge.comPut(handle, 'DisplayAlerts', false);
                    bridge.comCallWith(handle, 'Quit', []);
                }
                else {
                    bridge.comCallWith(handle, 'Clear', [1]);
                    bridge.comCallWith(handle, 'Quit', []);
                }
                apps[msg.app] = undefined;
            }
            return {
                type: 'status',
                excel: !!apps.excel,
                hwp: !!apps.hwp,
            };
        }
        case 'execute': {
            const hasDocument = !!apps.excel || !!apps.hwp;
            const result = hasDocument
                ? (0, executor_1.executeWithSavepoint)(msg.code, apps)
                : (0, executor_1.executeCode)(msg.code, apps);
            if (result.success) {
                return {
                    type: 'result',
                    id: msg.id,
                    success: true,
                    result: result.result,
                    logs: result.logs,
                };
            }
            else {
                return {
                    type: 'error',
                    id: msg.id,
                    success: false,
                    error: result.error || 'Unknown error',
                    stack: result.stack,
                    line: result.line,
                    code: result.code,
                    logs: result.logs,
                    restored: result.restored,
                };
            }
        }
        case 'list-members': {
            const handle = apps[msg.app];
            if (!handle) {
                return {
                    type: 'error',
                    id: msg.id,
                    success: false,
                    error: `${msg.app} is not running`,
                    code: '',
                    logs: [],
                    restored: false,
                };
            }
            const members = bridge.comListMembers(handle);
            return {
                type: 'members',
                id: msg.id,
                members: members.map((m) => ({
                    name: m.name,
                    kind: m.kind,
                    params: m.params,
                    returnType: m.returnType,
                })),
            };
        }
    }
}
// IPC: Electron utilityProcess에서는 parentPort 사용
if (typeof process.send === 'function') {
    // child_process.fork 모드 (테스트용)
    process.on('message', (msg) => {
        const response = handleMessage(msg);
        process.send(response);
    });
}
else if (process.parentPort) {
    // Electron utilityProcess 모드
    const { parentPort } = process;
    parentPort.on('message', (event) => {
        const response = handleMessage(event.data);
        parentPort.postMessage(response);
    });
}
else {
    // 단독 실행 모드 (REPL 테스트)
    console.log('COM Worker standalone mode');
    console.log('Apps:', Object.keys(apps).filter(k => !!apps[k]));
    const readline = require('readline');
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.on('line', (line) => {
        const trimmed = line.trim();
        if (!trimmed)
            return;
        if (trimmed === 'quit' || trimmed === 'exit') {
            // 앱 정리
            if (apps.excel)
                handleMessage({ type: 'quit', id: '0', app: 'excel' });
            if (apps.hwp)
                handleMessage({ type: 'quit', id: '0', app: 'hwp' });
            rl.close();
            process.exit(0);
        }
        if (trimmed.startsWith('launch ')) {
            const app = trimmed.split(' ')[1];
            const res = handleMessage({ type: 'launch', id: '0', app });
            console.log(JSON.stringify(res, null, 2));
            return;
        }
        // 나머지는 코드로 실행
        const res = handleMessage({ type: 'execute', id: '0', code: trimmed });
        console.log(JSON.stringify(res, null, 2));
    });
}
