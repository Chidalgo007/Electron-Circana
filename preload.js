const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  // config
  selectFolder: () => ipcRenderer.invoke("dialog:selectFolder"),
  selectFile: () => ipcRenderer.invoke("dialog:selectFile"),
  getConfig: () => ipcRenderer.invoke("config:get"),
  setConfig: (data) => ipcRenderer.invoke("config:set", data),
  // circana dashboard
  runAutomation: (runExcel) =>
    ipcRenderer.invoke("run-automation", { runExcel }),
  runAutomationPart2: (runExcel) =>
    ipcRenderer.invoke("run-automation-part2", { runExcel }),
  // circana pivot
  runExcel: (excelPath) => ipcRenderer.invoke("run-excel", excelPath),
  stopAutomation: () => ipcRenderer.invoke("stop-automation"),
  // NPD
  runNPD: (npdPath) => ipcRenderer.invoke("run-npd", npdPath),
  stopNPD: () => ipcRenderer.invoke("stop-npd"),
  // circana schedule
  setSchedule: (scheduleData) =>
    ipcRenderer.invoke("schedule:set", scheduleData),
  getSchedule: () => ipcRenderer.invoke("schedule:get"),
  cancelSchedule: () => ipcRenderer.invoke("schedule:cancel"),
  onRunSchedule: (callback) => ipcRenderer.on("run-schedule", callback),
  // log
  onLog: (callback) => ipcRenderer.on("log", callback),
  // timer
  onNpdTimer: (callback) => ipcRenderer.on("npd-timer", callback),
  onCircanaTimer: (callback) => ipcRenderer.on("circana-timer", callback),
});
