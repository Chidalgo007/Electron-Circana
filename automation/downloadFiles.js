// downloadFiles.js
const { chromium } = require("playwright-core");
const path = require("path");
const fs = require("fs");
const {
  loginUnify,
  downloadFromNotifications,
  moveTargetFiles,
  cleanupUuidFiles,
  ensureDir,
  TARGET_FILES,
} = require("./unifyHelpers");

(async () => {
  try {
    console.log("[downloadFiles] Starting...");

    // args: node downloadFiles.js <downloadDir> <destinationDir> <username> <password>
    const downloadDir = process.argv[2];
    const destinationDir = process.argv[3];
    const username = process.argv[4];
    const password = process.argv[5];

    if (!downloadDir || !destinationDir || !username || !password) {
      console.error(
        "Missing required arguments: downloadDir, destinationDir, username, password"
      );
      process.exit(1);
    }

    ensureDir(downloadDir);
    ensureDir(destinationDir);

    const userDataDir = path.resolve("./edge_profile");

    // find Chrome / Edge executable
    const chromePaths = [
      "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
      "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
      path.join(
        process.env.LOCALAPPDATA || "",
        "Google\\Chrome\\Application\\chrome.exe"
      ),
    ];
    const chromeExecutable = chromePaths.find((p) => fs.existsSync(p));
    if (!chromeExecutable) throw new Error("Chrome executable not found.");

    // launch persistent context
    const browser = await chromium.launchPersistentContext(userDataDir, {
      headless: false,
      executablePath: chromeExecutable,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
      acceptDownloads: true,
      channel: "msedge",
    });

    const page = browser.pages()[0];

    // login
    await loginUnify(page, { username, password });

    // download notifications
    const { successes, failures } = await downloadFromNotifications(
      page,
      downloadDir,
      { retries: 2, parallel: true }
    );
    console.log(
      `Downloads finished. Success: ${successes.length}, Failures: ${failures.length}`
    );

    // move files to final destination
    moveTargetFiles(downloadDir, destinationDir, TARGET_FILES);

    // cleanup temp UUID files
    cleanupUuidFiles(downloadDir);

    console.log("[downloadFiles] Automation completed successfully.");
    await browser.close();
    process.exit(0);
  } catch (err) {
    console.error("[downloadFiles] Automation failed:", err);
    process.exit(1);
  }
})();
