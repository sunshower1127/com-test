/**
 * Worker process entry point
 *
 * 실행 방법:
 * - Electron utilityProcess: main에서 fork
 * - 단독 테스트: node dist/worker/index.js
 */

import * as path from 'path';
import { initBridge } from './proxy';
import { initExecutor, executeWithSavepoint, executeCode } from './executor';
import type { WorkerRequest, WorkerResponse } from '../shared/types';

// .node 네이티브 모듈 로드
const nativePath = path.resolve(__dirname, '../../native/com_bridge_node.node');
let bridge: any;

try {
  bridge = require(nativePath);
} catch {
  // fallback: 빌드 output에서 로드
  const fallbackPath = path.resolve(__dirname, '../../../target/debug/com_bridge_node.node');
  bridge = require(fallbackPath);
}

// COM 초기화
bridge.comInit();

// Proxy & Executor에 bridge 주입
initBridge(bridge);
initExecutor(bridge);

// 앱 핸들 저장소
const apps: { excel?: unknown; hwp?: unknown } = {};

function handleMessage(msg: WorkerRequest): WorkerResponse {
  switch (msg.type) {
    case 'launch': {
      const progId = msg.app === 'excel' ? 'Excel.Application' : 'HWPFrame.HwpObject';
      const handle = bridge.comCreate(progId);
      apps[msg.app] = handle;

      // visible 설정
      if (msg.app === 'excel') {
        bridge.comPut(handle, 'Visible', true);
      } else {
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
        } else {
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
        ? executeWithSavepoint(msg.code, apps)
        : executeCode(msg.code, apps);

      if (result.success) {
        return {
          type: 'result',
          id: msg.id,
          success: true,
          result: result.result,
          logs: result.logs,
        };
      } else {
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
        members: members.map((m: any) => ({
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
  process.on('message', (msg: WorkerRequest) => {
    const response = handleMessage(msg);
    process.send!(response);
  });
} else if ((process as any).parentPort) {
  // Electron utilityProcess 모드
  const { parentPort } = process as any;
  parentPort.on('message', (event: { data: WorkerRequest }) => {
    const response = handleMessage(event.data);
    parentPort.postMessage(response);
  });
} else {
  // 단독 실행 모드 (REPL 테스트)
  console.log('COM Worker standalone mode');
  console.log('Apps:', Object.keys(apps).filter(k => !!(apps as any)[k]));

  const readline = require('readline');
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  rl.on('line', (line: string) => {
    const trimmed = line.trim();
    if (!trimmed) return;

    if (trimmed === 'quit' || trimmed === 'exit') {
      // 앱 정리
      if (apps.excel) handleMessage({ type: 'quit', id: '0', app: 'excel' });
      if (apps.hwp) handleMessage({ type: 'quit', id: '0', app: 'hwp' });
      rl.close();
      process.exit(0);
    }

    if (trimmed.startsWith('launch ')) {
      const app = trimmed.split(' ')[1] as 'excel' | 'hwp';
      const res = handleMessage({ type: 'launch', id: '0', app });
      console.log(JSON.stringify(res, null, 2));
      return;
    }

    // 나머지는 코드로 실행
    const res = handleMessage({ type: 'execute', id: '0', code: trimmed });
    console.log(JSON.stringify(res, null, 2));
  });
}
