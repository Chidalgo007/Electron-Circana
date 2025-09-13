const saveBtn = document.getElementById("saveConfigBtn");

// Download folder
document
  .getElementById("browseDownload")
  .addEventListener("click", async () => {
    const folderPath = await window.electronAPI.selectFolder();
    if (folderPath) {
      document.getElementById("downloadInput").value = folderPath;
    }
  });

// Destination folder
document
  .getElementById("browseDestination")
  .addEventListener("click", async () => {
    const folderPath = await window.electronAPI.selectFolder();
    if (folderPath) {
      document.getElementById("destinationInput").value = folderPath;
    }
  });

// Excel file
document.getElementById("browseExcel").addEventListener("click", async () => {
  const filePath = await window.electronAPI.selectFile();
  if (filePath) {
    document.getElementById("excelInput").value = filePath;
  }
});

// NPD file
document.getElementById("browseNPD").addEventListener("click", async () => {
  const filePath = await window.electronAPI.selectFile();
  if (filePath) {
    document.getElementById("npdInput").value = filePath;
  }
});

// configuration credentials
let config = null;
async function loadConfig() {
  config = await window.electronAPI.getConfig();
  document.getElementById("username").value = config.username || "";
  document.getElementById("password").value = config.password || "";
  document.getElementById("downloadInput").value = config.downloadPath || "";
  document.getElementById("destinationInput").value =
    config.destinationPath || "";
  document.getElementById("excelInput").value = config.excelPath || "";
  document.getElementById("npdInput").value = config.npdPath || "";
}

saveBtn.addEventListener("click", async () => {
  const data = {
    username: document.getElementById("username").value,
    password: document.getElementById("password").value,
    downloadPath: document.getElementById("downloadInput").value,
    destinationPath: document.getElementById("destinationInput").value,
    excelPath: document.getElementById("excelInput").value,
    npdPath: document.getElementById("npdInput").value,
  };
  await window.electronAPI.setConfig(data);
  alert("Config saved ✅");
});

loadConfig();
// ------- end config -------

// -------- schedule ---------
// set min datetime to current datetime
document.addEventListener("DOMContentLoaded", () => {
  const scheduleInput = document.getElementById("scheduleInput");
  const now = new Date().toISOString().slice(0, 16); // yyyy-MM-ddThh:mm
  scheduleInput.min = now;

  scheduleInput.addEventListener("input", () => {
    const selected = new Date(scheduleInput.value);
    const current = new Date();
    if (selected < current) {
      scheduleInput.value = " "; // reset to current datetime
      alert("Please select a future date and time!");
    }
  });
});

async function schedule() {
  const dateInput = document.getElementById("scheduleInput");
  const repeat = document.getElementById("repeatSchedule").checked;
  const selectedType = document.querySelector(
    'input[name="scheduleType"]:checked'
  );

  // make sure date and time are set
  if (!dateInput.value) {
    alert("Please pick a valid date/time.");
    return;
  }

  const scheduleData = {
    date: dateInput.value,
    repeat,
    type: selectedType?.id || "scheduleCircanaDashboardExcel", // fallback to both
  };

  await window.electronAPI.setSchedule(scheduleData);
  await renderSchedule();
  if (result.success) {
    await renderSchedule();
    alert("Schedule set ✅");
  } else {
    alert("⚠️ Schedule NOT saved. Please complete all configuration values.");
  }
}

async function clearSchedule() {
  await window.electronAPI.cancelSchedule();
  await renderSchedule();
  alert("Schedule cleared ✅");
}

// Render schedule into <pre>
async function renderSchedule() {
  const schedules = await window.electronAPI.getSchedule();
  const viewEl = document.getElementById("viewSchedule");

  if (schedules.length) {
    viewEl.textContent = schedules
      .map((s) => {
        const repeatText = s.repeat ? "🔁 Weekly" : "📅 One-time";
        let typeText = "❓ Unknown";
        if (s.type === "scheduleCircanaPivot") typeText = "Circana Pivot Only";
        if (s.type === "scheduleCircanaDashboard")
          typeText = "Circana Dashboard Only";
        if (s.type === "scheduleCircanaDashboardExcel")
          typeText = "Circana Dashboard + Circana Pivot";

        return `- ${repeatText} | ${typeText} | At: ${s.date}`;
      })
      .join("\n");
  } else {
    viewEl.textContent = "❌ No schedules set";
  }
}

// When main tells us to run scheduled automation
window.electronAPI.onRunSchedule((event, type) => {
  if (type === "pivot") {
    window.electronAPI.runExcel(config.excelPath);
  } else if (type === "dashboard") {
    window.electronAPI.runAutomation(false);
  } else if (type === "both") {
    window.electronAPI.runAutomation(true);
  }
});

// On load, show current schedule
document.addEventListener("DOMContentLoaded", renderSchedule);
