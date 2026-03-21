import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('comBridge', {
  send: (msg: any) => ipcRenderer.invoke('send-to-worker', msg),
  onResponse: (callback: (msg: any) => void) => {
    ipcRenderer.on('worker-response', (_event, msg) => callback(msg));
  },
});
