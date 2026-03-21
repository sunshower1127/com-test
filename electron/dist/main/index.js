"use strict";
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
const electron_1 = require("electron");
const child_process_1 = require("child_process");
const path = __importStar(require("path"));
let mainWindow = null;
let worker = null;
function createWindow() {
    mainWindow = new electron_1.BrowserWindow({
        width: 900,
        height: 700,
        title: 'COM Bridge',
        webPreferences: {
            preload: path.join(__dirname, '../renderer/preload.js'),
            contextIsolation: true,
            nodeIntegration: false,
        },
    });
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));
}
function startWorker() {
    const workerPath = path.join(__dirname, '../worker/index.js');
    worker = (0, child_process_1.fork)(workerPath, [], { stdio: ['pipe', 'pipe', 'pipe', 'ipc'] });
    worker.on('message', (msg) => {
        mainWindow?.webContents.send('worker-response', msg);
    });
    worker.stdout?.on('data', (data) => {
        console.log('[worker stdout]', data.toString());
    });
    worker.stderr?.on('data', (data) => {
        console.error('[worker stderr]', data.toString());
    });
    worker.on('exit', (code) => {
        console.log(`Worker exited with code ${code}`);
        worker = null;
    });
}
electron_1.app.whenReady().then(() => {
    startWorker();
    createWindow();
});
// Renderer → Main → Worker
electron_1.ipcMain.handle('send-to-worker', (_event, msg) => {
    if (worker) {
        worker.send(msg);
    }
    else {
        mainWindow?.webContents.send('worker-response', {
            type: 'error',
            id: msg.id || '0',
            success: false,
            error: 'Worker not running',
            code: '',
            logs: [],
            restored: false,
        });
    }
});
electron_1.app.on('window-all-closed', () => {
    if (worker) {
        worker.send({ type: 'quit', id: '0', app: 'excel' });
        worker.send({ type: 'quit', id: '0', app: 'hwp' });
        setTimeout(() => worker?.kill(), 2000);
    }
    electron_1.app.quit();
});
