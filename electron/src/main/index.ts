import { app, BrowserWindow, ipcMain } from 'electron';
import { fork, ChildProcess } from 'child_process';
import * as path from 'path';

let mainWindow: BrowserWindow | null = null;
let worker: ChildProcess | null = null;

function createWindow() {
  mainWindow = new BrowserWindow({
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
  worker = fork(workerPath, [], { stdio: ['pipe', 'pipe', 'pipe', 'ipc'] });

  worker.on('message', (msg: any) => {
    mainWindow?.webContents.send('worker-response', msg);
  });

  worker.stdout?.on('data', (data: Buffer) => {
    console.log('[worker stdout]', data.toString());
  });

  worker.stderr?.on('data', (data: Buffer) => {
    console.error('[worker stderr]', data.toString());
  });

  worker.on('exit', (code) => {
    console.log(`Worker exited with code ${code}`);
    worker = null;
  });
}

app.whenReady().then(() => {
  startWorker();
  createWindow();
});

// Renderer → Main → Worker
ipcMain.handle('send-to-worker', (_event, msg) => {
  if (worker) {
    worker.send(msg);
  } else {
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

app.on('window-all-closed', () => {
  if (worker) {
    worker.send({ type: 'quit', id: '0', app: 'excel' });
    worker.send({ type: 'quit', id: '0', app: 'hwp' });
    setTimeout(() => worker?.kill(), 2000);
  }
  app.quit();
});
