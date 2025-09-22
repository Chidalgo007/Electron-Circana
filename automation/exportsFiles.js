// exportFiles.js
const { chromium } = require("playwright-core");
const path = require("path");
const fs = require("fs");
const {
  loginUnify,
  navigateAndQueueExports,
  ensureDir,
} = require("./unifyHelpers");

(async () => {
  try {
    console.log("[exportFiles] Starting export queue...");

    // args: node exportFiles.js <username> <password>
    const username = process.argv[4];
    const password = process.argv[5];

    if (!username || !password) {
      console.error("Missing username/password arguments");
      process.exit(1);
    }

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
      channel: "msedge",
    });

    const page = browser.pages()[0];

    // login
    await loginUnify(page, { username, password });

    // navigate and queue all exports
    await navigateAndQueueExports(page);

    console.log(
      "[exportFiles] All exports queued. You can now close the browser or start download process."
    );

    await browser.close();
    process.exit(0);
  } catch (err) {
    console.error("[exportFiles] Automation failed:", err);
    process.exit(1);
  }
})();
