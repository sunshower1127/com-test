/**
 * VM 샌드박스 실행기 + 세이브포인트 자동 롤백
 */

import * as vm from 'vm';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createComProxy } from './proxy';

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

let bridge: ComBridge;

export function initExecutor(b: ComBridge) {
  bridge = b;
}

/**
 * LLM 코드를 VM 샌드박스에서 실행
 */
function runInSandbox(code: string, apps: AppHandles): ExecutionResult {
  const logs: string[] = [];

  const sandbox: Record<string, unknown> = {
    console: {
      log: (...args: unknown[]) => logs.push(args.map(String).join(' ')),
      error: (...args: unknown[]) => logs.push('[ERROR] ' + args.map(String).join(' ')),
    },
    result: null,
  };

  // COM proxy 객체를 sandbox에 주입
  if (apps.excel) {
    sandbox.excel = createComProxy(apps.excel);
  }
  if (apps.hwp) {
    sandbox.hwp = createComProxy(apps.hwp);
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
  } catch (e: unknown) {
    const err = e as Error;
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
export function executeWithSavepoint(
  code: string,
  apps: AppHandles,
): ExecutionResult {
  const tempDir = os.tmpdir();
  const timestamp = Date.now();
  let savepointPath: string | null = null;
  let activeApp: 'excel' | 'hwp' | null = null;

  // 1. 세이브포인트 생성
  try {
    if (apps.excel) {
      savepointPath = path.join(tempDir, `savepoint_excel_${timestamp}.xlsx`);
      bridge.comCallWith(apps.excel, 'SaveCopyAs', [savepointPath]);
      activeApp = 'excel';
    } else if (apps.hwp) {
      savepointPath = path.join(tempDir, `savepoint_hwp_${timestamp}.hwp`);
      bridge.comCallWith(apps.hwp, 'SaveAs', [savepointPath, 'HWP', '']);
      activeApp = 'hwp';
    }
  } catch {
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
      } else if (activeApp === 'hwp' && apps.hwp) {
        bridge.comCallWith(apps.hwp, 'Clear', [1]);
        bridge.comCallWith(apps.hwp, 'Open', [savepointPath]);
      }
      result.restored = true;
    } catch {
      // 복원 실패 — 에러에 기록
      result.error = (result.error || '') + ' [WARN: savepoint restore failed]';
    }

    // 임시파일 정리
    try { fs.unlinkSync(savepointPath); } catch { /* ignore */ }
  } else if (savepointPath) {
    // 성공 시 임시파일 정리
    try { fs.unlinkSync(savepointPath); } catch { /* ignore */ }
  }

  return result;
}

/**
 * 세이브포인트 없이 실행 (새 문서 등)
 */
export function executeCode(code: string, apps: AppHandles): ExecutionResult {
  return runInSandbox(code, apps);
}
