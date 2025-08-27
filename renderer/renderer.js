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

// configuration credentials
async function loadConfig() {
  const config = await window.electronAPI.getConfig();
  document.getElementById("username").value = config.username;
  document.getElementById("password").value = config.password;
  document.getElementById("downloadInput").value = config.downloadPath;
  document.getElementById("destinationInput").value = config.destinationPath;
  document.getElementById("excelInput").value = config.excelPath;
}

saveBtn.addEventListener("click", async () => {
  const data = {
    username: document.getElementById("username").value,
    password: document.getElementById("password").value,
    downloadPath: document.getElementById("downloadInput").value,
    destinationPath: document.getElementById("destinationInput").value,
    excelPath: document.getElementById("excelInput").value,
  };
  await window.electronAPI.setConfig(data);
  alert("Config saved ✅");
});

loadConfig();
// ------- end config -------
