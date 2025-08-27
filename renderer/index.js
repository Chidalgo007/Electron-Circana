const logBox = document.getElementById("logBox");
// check box to set to run the Circana Pivot automation
function runAutomation() {
  const runExcel = document.getElementById("linkProcess").checked;
  window.electronAPI.runAutomation(runExcel);
}

function runExcel() {
  const filePath = document.getElementById("excelInput").value;
  if (!filePath) {
    alert("Please specify the Circana Pivot Excel file path in configuration.");
    return;
  }
  window.electronAPI.runExcel(filePath);
}
// stop the automation process
async function stopCircanaPivot() {
  const killed = await window.electronAPI.stopAutomation();
  if (killed) {
    logBox.textContent += "🛑 Circana Pivot process killed.\n";
  }
}

// log output
window.electronAPI.onLog((event, msg) => {
  logBox.textContent += msg + "\n";
  // Auto-scroll to bottom
  logBox.scrollTop = logBox.scrollHeight;
});

// change the tab content
function openTab(tabId) {
  document
    .querySelectorAll(".tabcontent")
    .forEach((el) => (el.style.display = "none"));
  document.getElementById(tabId).style.display = "flex";

  logBox.style.display = ["config", "schedule"].includes(tabId)
    ? "none"
    : "block";

  document.querySelectorAll(".tabs button").forEach((btn) => {
    btn.classList.remove("active");
    if (btn.getAttribute("tabid") === tabId) btn.classList.add("active");
  });
}

openTab("Circana Dashboard"); // default
