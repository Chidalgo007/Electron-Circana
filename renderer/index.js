const logBox = document.getElementById("logBox");
const logBoxBtn = document.getElementById("clearLog");
const circanaTime = document.getElementById("circanaTime");
const npdTime = document.getElementById("npdTime");

// check box to set to run the Circana Pivot automation
function runAutomation() {
  const runExcel = document.getElementById("linkProcess").checked;
  window.electronAPI.runAutomation(runExcel);
}
function runAutomationPartII() {
  const runExcel = document.getElementById("linkProcess").checked;
  window.electronAPI.runAutomationPart2(runExcel);
}

function runExcel() {
  const filePath = document.getElementById("excelInput").value;
  if (!filePath) {
    alert("Please specify the Circana Pivot Excel file path in configuration.");
    return;
  }
  circanaTime.textContent = "00:00:00"; // reset timer
  window.electronAPI.runExcel(filePath);
}
// stop the automation process
async function stopCircanaPivot() {
  const killed = await window.electronAPI.stopAutomation();
  if (killed) {
    logBox.textContent += "🛑 Circana Pivot process killed.\n";
  }
  setTimeout(() => {
    circanaTime.textContent = "--:--:--"; // reset timer display
  }, 2000);
}

// Timer for circana dashboard
window.electronAPI.onCircanaTimer((event, elapsed) => {
  const hours = Math.floor(elapsed / 3600);
  const mins = Math.floor((elapsed % 3600) / 60);
  const secs = elapsed % 60;

  circanaTime.textContent = `${hours.toString().padStart(2, "0")}:${mins
    .toString()
    .padStart(2, "0")}:${secs.toString().padStart(2, "0")}`;
});

// run the NPD process
function runNPD() {
  const filePath = document.getElementById("npdInput").value;
  if (!filePath) {
    alert("Please specify the NPD Excel file path in configuration.");
    return;
  }
  npdTime.textContent = "00:00:00"; // reset timer
  window.electronAPI.runNPD(filePath);
}

// stop the automation process
async function stopNPD() {
  const killed = await window.electronAPI.stopNPD();
  if (killed) {
    logBox.textContent += "🛑 NPD process killed.\n";
  }
  setTimeout(() => {
    npdTime.textContent = "--:--:--"; // reset timer display
  }, 2000);
}

// Timer for NPD
window.electronAPI.onNpdTimer((event, elapsed) => {
  const hours = Math.floor(elapsed / 3600);
  const mins = Math.floor((elapsed % 3600) / 60);
  const secs = elapsed % 60;

  npdTime.textContent = `${hours.toString().padStart(2, "0")}:${mins
    .toString()
    .padStart(2, "0")}:${secs.toString().padStart(2, "0")}`;
});

// log output
window.electronAPI.onLog((event, msg) => {
  logBox.textContent += msg + "\n";
  // Auto-scroll to bottom
  logBox.scrollTop = logBox.scrollHeight;
});

function clearLog() {
  logBox.textContent = "🗑️ Content cleared...\n";
}

// change the tab content
function openTab(tabId) {
  document
    .querySelectorAll(".tabcontent")
    .forEach((el) => (el.style.display = "none"));
  document.getElementById(tabId).style.display = "flex";

  logBox.style.display = ["config", "schedule"].includes(tabId)
    ? "none"
    : "block";

  logBoxBtn.style.display = ["config", "schedule"].includes(tabId)
    ? "none"
    : "block";

  document.querySelectorAll(".tabs button").forEach((btn) => {
    btn.classList.remove("active");
    if (btn.getAttribute("tabid") === tabId) btn.classList.add("active");
  });
}

openTab("Circana-Dashboard"); // default
