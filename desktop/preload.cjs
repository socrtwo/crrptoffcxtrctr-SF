const { contextBridge, ipcRenderer } = require("electron");
const fs = require("node:fs");

contextBridge.exposeInMainWorld("electronBridge", {
  onOpenFile: (cb) => ipcRenderer.on("open-file", (_e, path) => cb(path)),
  readFile: (path) => fs.promises.readFile(path),
});
