"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const electron_1 = require("electron");
electron_1.contextBridge.exposeInMainWorld('comBridge', {
    send: (msg) => electron_1.ipcRenderer.invoke('send-to-worker', msg),
    onResponse: (callback) => {
        electron_1.ipcRenderer.on('worker-response', (_event, msg) => callback(msg));
    },
});
