import { app, BrowserWindow, dialog, ipcMain } from "electron";
import { spawn, exec } from "child_process";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import store from "./config.js";
import { saveCredential, getCredential } from "./secureStore.js";

let mainWindow;
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 500,
    height: 500,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
    },
  });

  mainWindow.loadFile("renderer/index.html");
}

app.whenReady().then(createWindow);

// Exit on all windows closed
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

// =========== Functions ================
let circanaPivot;
let excelPID;
let canceled = false;
function runPivotExcel(filePath) {
  return new Promise((resolve, reject) => {
    const exePath = path.join(__dirname, "python", "CircanaElectron.exe");

    circanaPivot = spawn(exePath, [filePath], {
      stdio: "pipe",
      windowsHide: true,
    });

    circanaPivot.stdout.on("data", (data) => {
      const text = data.toString();
      if (!text.includes("The RPC server is unavailable."))
        mainWindow.webContents.send("log", text);
      const match = text.match(/Excel started with PID:(\d+)/);
      if (match) excelPID = parseInt(match[1], 10);
      // Ignore RPC server unavailable messages
    });

    circanaPivot.stderr.on("data", (data) => {
      const text = data.toString();
      if (!text.includes("The RPC server is unavailable."))
        mainWindow.webContents.send("log", `[ERROR] ${text}`);
    });

    circanaPivot.on("close", (code) => {
      if (code === 0) {
        mainWindow.webContents.send(
          "log",
          "✅ Pivot automation finished successfully."
        );
        resolve(true);
      } else {
        if (!canceled)
          mainWindow.webContents.send(
            "log",
            `❌ Pivot automation failed with code ${code}`
          );
        resolve(false);
      }
    });

    circanaPivot.on("error", (err) => {
      mainWindow.webContents.send(
        "log",
        `❌ Failed to start process: ${err.message}`
      );
      reject(err);
    });
  });
}

// ========= IPCs ================

// IPC - circana pivot excel run ==============
ipcMain.handle("run-excel", () => {
  const excelPath = store.get("excelPath"); // Excel file path from config
  if (!excelPath) {
    mainWindow.webContents.send("log", "[Pivot] No Excel path configured!");
    return;
  }
  runPivotExcel(excelPath);
});

// IPC: ==== stop automation =====================
ipcMain.handle("stop-automation", () => {
  if (!circanaPivot && !excelPID) return; // nothing to stop
  // 1️⃣ Kill Python process immediately
  circanaPivot.kill(); // default SIGTERM
  mainWindow.webContents.send("log", "🛑 Python process killed.");

  // 2️⃣ Kill Excel PID if known
  setTimeout(() => {
    if (excelPID) {
      exec(`taskkill /PID ${excelPID} /F`, (err) => {
        if (!err) {
          mainWindow.webContents.send(
            "log",
            `🛑 Excel PID ${excelPID} killed.`
          );
          canceled = true;
        } else {
          mainWindow.webContents.send(
            "log",
            `[ERROR] Could not kill Excel PID ${excelPID}: ${err.message}`
          );
        }
        excelPID = null; // reset
      });
    }
  }, 500); // 500ms delay

  // 3️⃣ Clear Python process reference
  circanaPivot = null;
});

// IPC: ==== run website automation download file =====================
ipcMain.handle("run-automation", async (event, { runExcel }) => {
  return new Promise((resolve) => {
    const automation = spawn("node", [
      path.join(__dirname, "automation/download.js"),
    ]);

    automation.stdout.on("data", (data) => {
      mainWindow.webContents.send("log", `[Automation] ${data.toString()}`);
    });

    automation.stderr.on("data", (data) => {
      mainWindow.webContents.send(
        "log",
        `[Automation ERROR] ${data.toString()}`
      );
    });

    automation.on("close", (code) => {
      if (code === 0 && runExcel) {
        const excelPath = store.get("excelPath"); // Excel file path from config
        if (!excelPath) {
          mainWindow.webContents.send(
            "log",
            "[Pivot] No Excel path configured!"
          );
          return;
        }
        runPivotExcel(excelPath);
      }
      resolve({ success: code === 0 });
    });
  });
});

// ============ configuration tab part ========================
// Load config credentials for UI
ipcMain.handle("config:get", async () => {
  return {
    downloadPath: store.get("downloadPath"),
    destinationPath: store.get("destinationPath"),
    excelPath: store.get("excelPath"),
    username: (await getCredential("circana-username")) || "",
    password: (await getCredential("circana-password")) || "",
  };
});

// Save config credentials from UI
ipcMain.handle("config:set", async (event, data) => {
  store.set("downloadPath", data.downloadPath);
  store.set("destinationPath", data.destinationPath);
  store.set("excelPath", data.excelPath);

  if (data.username) await saveCredential("circana-username", data.username);
  if (data.password) await saveCredential("circana-password", data.password);

  return true;
});

// Folder picker
ipcMain.handle("dialog:selectFolder", async () => {
  console.log("Opening folder dialog…");
  const result = await dialog.showOpenDialog({
    properties: ["openDirectory"],
  });
  return result.canceled ? null : result.filePaths[0];
});

// Excel file picker
ipcMain.handle("dialog:selectFile", async () => {
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xlsm", "xls"] }],
  });
  return result.canceled ? null : result.filePaths[0];
});
