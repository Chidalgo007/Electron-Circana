const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  selectFolder: () => ipcRenderer.invoke("dialog:selectFolder"),
  selectFile: () => ipcRenderer.invoke("dialog:selectFile"),
  getConfig: () => ipcRenderer.invoke("config:get"),
  setConfig: (data) => ipcRenderer.invoke("config:set", data),
  runAutomation: (runExcel) =>
    ipcRenderer.invoke("run-automation", { runExcel }),
  runExcel: (excelPath) => ipcRenderer.invoke("run-excel", excelPath),
  stopAutomation: () => ipcRenderer.invoke("stop-automation"),
  onLog: (callback) => ipcRenderer.on("log", callback),
});
