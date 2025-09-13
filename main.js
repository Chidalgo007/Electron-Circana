const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const { spawn, exec, fork } = require("child_process");
const path = require("path");
const store = require("./config.js");
const { saveCredential, getCredential } = require("./secureStore.js");

// Disable GPU and cache (fixes "Unable to move cache" errors)
app.commandLine.appendSwitch("disable-gpu");
app.commandLine.appendSwitch("disable-software-rasterizer");
app.commandLine.appendSwitch("disable-gpu-shader-disk-cache");
app.commandLine.appendSwitch("disable-gpu-program-cache");

let mainWindow;

// Single instance lock
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  app.on("second-instance", () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 560,
    height: 500,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
    },
    resizable: false,
    autoHideMenuBar: true,
  });

  mainWindow.loadFile("renderer/index.html");
}

// App ready
app.whenReady().then(() => {
  createWindow();
  startScheduler();
});

// Exit on all windows closed
app.on("window-all-closed", () => app.quit());

// ================= Automation Functions =================
let circanaPivot;
let excelPID;
let canceled = false;

function runPivotExcel(filePath) {
  return new Promise((resolve, reject) => {
    const execPath = app.isPackaged
      ? path.join(
          process.resourcesPath,
          "app.asar.unpacked",
          "python",
          "CircanaElectron.exe"
        )
      : path.join(__dirname, "python", "CircanaElectron.exe");

    if (!execPath || !filePath) {
      mainWindow.webContents.send(
        "log",
        "[ERROR] Missing executable or file path."
      );
      return reject(new Error("Missing executable or file path."));
    }

    circanaPivot = spawn(execPath, [filePath], {
      stdio: "pipe",
      windowsHide: true,
    });
    //timer for elapsed time
    const startTime = Date.now();
    let timer = setInterval(() => {
      const elapsed = Math.floor((Date.now() - startTime) / 1000);
      mainWindow.webContents.send("circana-timer", elapsed);
    }, 1000);

    circanaPivot.stdout.on("data", (data) => {
      const text = data.toString();
      if (!text.includes("The RPC server is unavailable."))
        mainWindow.webContents.send("log", text);
      const match = text.match(/Excel started with PID:(\d+)/);
      if (match) excelPID = parseInt(match[1], 10);
    });

    circanaPivot.stderr.on("data", (data) => {
      const text = data.toString();
      if (!text.includes("The RPC server is unavailable."))
        mainWindow.webContents.send("log", `[ERROR] ${text}`);
    });

    circanaPivot.on("close", (code) => {
      clearInterval(timer); // 🛑 Stop timer
      if (code === 0)
        mainWindow.webContents.send(
          "log",
          "✅ Pivot automation finished successfully."
        );
      else if (!canceled)
        mainWindow.webContents.send(
          "log",
          `❌ Pivot automation failed with code ${code}`
        );
      resolve(code === 0);
    });

    circanaPivot.on("error", (err) => {
      clearInterval(timer); // 🛑 Stop timer
      mainWindow.webContents.send(
        "log",
        `❌ Failed to start process: ${err.message}`
      );
      reject(err);
    });
  });
}

let NPD;
let npdPID;
let NPDcanceled = false;
const ignoredErrorsOnCancel = [
  "Pivot refreshing failed:",
  "Date filter update failed:",
  "Power BI Pivots failed:",
  "Critical error",
];

function runNPDProcess(filePath) {
  return new Promise((resolve, reject) => {
    const execPath = app.isPackaged
      ? path.join(
          process.resourcesPath,
          "app.asar.unpacked",
          "python",
          "NPD.exe"
        )
      : path.join(__dirname, "python", "NPD.exe");

    if (!execPath || !filePath) {
      mainWindow.webContents.send(
        "log",
        "[ERROR] Missing executable or file path."
      );
      return reject(new Error("Missing executable or file path."));
    }

    NPD = spawn(execPath, [filePath], {
      stdio: "pipe",
      windowsHide: true,
    });

    //timer for elapsed time
    const startTime = Date.now();
    let timer = setInterval(() => {
      const elapsed = Math.floor((Date.now() - startTime) / 1000);
      console.log("Elapsed:", elapsed); // Debug line
      mainWindow.webContents.send("npd-timer", elapsed);
    }, 1000);

    // message from npd process
    NPD.stdout.on("data", (data) => {
      const text = data.toString();
      // Suppress known errors if NPD was canceled
      if (
        NPDcanceled &&
        ignoredErrorsOnCancel.some((msg) => text.includes(msg))
      )
        return;
      if (!text.includes("The RPC server is unavailable."))
        mainWindow.webContents.send("log", text);
      const match = text.match(/Excel started with PID:(\d+)/);
      if (match) npdPID = parseInt(match[1], 10);
      console.log(`npdPID: ${npdPID}`);
    });

    NPD.stderr.on("data", (data) => {
      const text = data.toString();
      if (!text.includes("The RPC server is unavailable."))
        mainWindow.webContents.send("log", `[ERROR] ${text}`);
    });

    NPD.on("close", (code) => {
      clearInterval(timer); // 🛑 Stop timer
      if (code === 0)
        mainWindow.webContents.send(
          "log",
          "✅ NPD Refresh finished successfully."
        );
      else if (!NPDcanceled)
        mainWindow.webContents.send(
          "log",
          `❌ NPD Refresh failed with code ${code}`
        );
      resolve(code === 0);
    });

    NPD.on("error", (err) => {
      clearInterval(timer); // 🛑 Stop timer
      mainWindow.webContents.send(
        "log",
        `❌ Failed to start process: ${err.message}`
      );
      reject(err);
    });
  });
}

function runExcelDateUpdate() {
  return new Promise((resolve, reject) => {
    const flatFilesPath = store.get("destinationPath");
    if (!flatFilesPath) return reject(new Error("Destination path not set"));

    const execPath = app.isPackaged
      ? path.join(
          process.resourcesPath,
          "app.asar.unpacked",
          "python",
          "ExcelDateUpdate.exe"
        )
      : path.join(__dirname, "python", "ExcelDateUpdate.exe");

    const excelProcess = spawn(execPath, [flatFilesPath], {
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    });

    excelProcess.stdout.on("data", (data) => {
      mainWindow.webContents.send("log", data.toString().trim());
    });

    excelProcess.stderr.on("data", (data) => {
      mainWindow.webContents.send("log", `[ERROR] ${data.toString().trim()}`);
    });

    excelProcess.on("close", (code) => resolve(code === 0));

    excelProcess.on("error", (err) => reject(err));
  });
}

// ================= Schedule Functions =================
let scheduleTimer = null;

function checkSchedule() {
  const schedules = store.get("schedules") || [];
  const now = new Date();

  schedules.forEach((schedule, index) => {
    const target = new Date(schedule.date);
    const isMatch =
      (schedule.repeat &&
        now.getDay() === target.getDay() &&
        now.getHours() === target.getHours() &&
        now.getMinutes() === target.getMinutes()) ||
      (!schedule.repeat && Math.abs(now - target) < 60000);

    if (isMatch) {
      if (mainWindow) {
        mainWindow.show();
        mainWindow.focus();
      }

      switch (schedule.type) {
        case "scheduleCircanaPivot":
          mainWindow.webContents.send("run-schedule", "pivot");
          break;
        case "scheduleCircanaDashboard":
          mainWindow.webContents.send("run-schedule", "dashboard");
          break;
        case "scheduleCircanaDashboardExcel":
          mainWindow.webContents.send("run-schedule", "both");
          break;
      }

      if (!schedule.repeat) {
        schedules.splice(index, 1);
        store.set("schedules", schedules);
      }
    }
  });
}

function startScheduler() {
  if (scheduleTimer) clearInterval(scheduleTimer);
  scheduleTimer = setInterval(checkSchedule, 30 * 1000); // every 30s
}

// ================= IPC Handlers =================

// Run Pivot Excel
ipcMain.handle("run-excel", () => {
  const excelPath = store.get("excelPath");
  if (!excelPath)
    return mainWindow.webContents.send(
      "log",
      "[Pivot] No Excel path configured!"
    );
  runPivotExcel(excelPath);
});

// Stop excel Automation
ipcMain.handle("stop-automation", () => {
  if (circanaPivot) circanaPivot.kill();
  if (excelPID) {
    setTimeout(() => {
      exec(`taskkill /PID ${excelPID} /F`, () => {
        canceled = true;
        excelPID = null;
      });
    }, 500);
  }
  circanaPivot = null;
  mainWindow.webContents.send("log", "🛑 Automation Circana stopped.");
});
// Run NPD Excel
ipcMain.handle("run-npd", () => {
  const npdPath = store.get("npdPath");
  if (!npdPath)
    return mainWindow.webContents.send(
      "log",
      "[NPD] No Excel path configured!"
    );
  runNPDProcess(npdPath);
});

// Stop NPD Automation
ipcMain.handle("stop-npd", () => {
  if (NPD && !NPD.killed) NPD.kill(); // kill python
  // 2️⃣ Kill Excel immediately by PID
  if (npdPID) {
    setTimeout(() => {
      exec(`taskkill /PID ${npdPID} /F /T`, () => {
        NPDcanceled = true;
        npdPID = null; // clear PID after kill
      });
    }, 500);
  }

  mainWindow.webContents.send("log", "🛑 Automation NPD stopped.");
  NPD = null; // clear process reference
});

// Run website automation
ipcMain.handle("run-automation", async (event, { runExcel }) => {
  const downloadPath = store.get("downloadPath");
  const destinationPath = store.get("destinationPath");
  const username = await getCredential("circana-username");
  const password = await getCredential("circana-password");

  if (!downloadPath || !destinationPath || !username || !password) {
    mainWindow.webContents.send(
      "log",
      "[Automation] Missing configuration or credentials!"
    );
    return { success: false };
  }

  const automationPath = app.isPackaged
    ? path.join(
        process.resourcesPath,
        "app.asar.unpacked",
        "automation",
        "download.js"
      )
    : path.join(__dirname, "automation", "download.js");

  const automation = fork(
    automationPath,
    [downloadPath, destinationPath, username, password],
    {
      stdio: "pipe",
    }
  );

  automation.stdout.on("data", (data) =>
    mainWindow.webContents.send("log", `[Automation] ${data.toString()}`)
  );
  automation.stderr.on("data", (data) =>
    mainWindow.webContents.send("log", `[Automation ERROR] ${data.toString()}`)
  );

  return new Promise((resolve) => {
    automation.on("close", async (code) => {
      try {
        let result;
        if (code === 0) {
          event.reply("automation-done", { success: true });
          if (runExcel) {
            const excelPath = store.get("excelPath");
            if (excelPath) await runPivotExcel(excelPath);
          }
          await runExcelDateUpdate();
          result = { success: true };
        } else if (code === 2) {
          result = { success: false, reason: "User closed browser" };
        } else {
          result = { success: false, reason: "Automation failed" };
        }

        resolve(result);
      } catch (err) {
        console.error("Excel automation failed:", err);
        resolve({ success: false, error: err.message });
      }
    });
  });
});

// Config get
ipcMain.handle("config:get", async () => ({
  downloadPath: store.get("downloadPath"),
  destinationPath: store.get("destinationPath"),
  excelPath: store.get("excelPath"),
  npdPath: store.get("npdPath"),
  username: (await getCredential("circana-username")) || "",
  password: (await getCredential("circana-password")) || "",
}));

// Config set
ipcMain.handle("config:set", async (event, data) => {
  store.set("downloadPath", data.downloadPath);
  store.set("destinationPath", data.destinationPath);
  store.set("excelPath", data.excelPath);
  store.set("npdPath", data.npdPath);

  if (data.username) await saveCredential("circana-username", data.username);
  if (data.password) await saveCredential("circana-password", data.password);

  return true;
});

// Folder Picker
ipcMain.handle("dialog:selectFolder", async () => {
  const result = await dialog.showOpenDialog({ properties: ["openDirectory"] });
  return result.canceled ? null : result.filePaths[0];
});

// File Picker
ipcMain.handle("dialog:selectFile", async () => {
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xlsm", "xls"] }],
  });
  return result.canceled ? null : result.filePaths[0];
});

// ==== Schedule set/get/cancel ====
ipcMain.handle("schedule:set", async (event, data) => {
  const downloadPath = store.get("downloadPath");
  const destinationPath = store.get("destinationPath");
  const excelPath = store.get("excelPath");
  const username = await getCredential("circana-username");
  const password = await getCredential("circana-password");

  if (
    !downloadPath ||
    !destinationPath ||
    !excelPath ||
    !username ||
    !password
  ) {
    dialog.showErrorBox(
      "Missing Configuration",
      "Please set all required configuration values."
    );
    return { success: false };
  }

  const schedules = store.get("schedules") || [];
  schedules.push(data);
  store.set("schedules", schedules);
  startScheduler();
  return { success: true, saved: data };
});

ipcMain.handle("schedule:get", () => store.get("schedules") || []);

ipcMain.handle("schedule:cancel", () => {
  store.delete("schedules");
  if (scheduleTimer) clearInterval(scheduleTimer);
  return true;
});
