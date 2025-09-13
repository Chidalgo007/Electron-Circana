const { chromium } = require("playwright-core");
const fs = require("fs");
const path = require("path");

// ---------- Unify Automation ----------
async function loginUnify(page, { username, password }) {
  console.log("[loginUnify] Navigating to Unify...");

  const targetUrl = "https://unify.ap.iriworldwide.com/client1/index.html";
  let attempt = 0;
  while (attempt < 2) {
    try {
      attempt++;
      console.log(`[loginUnify] Attempt ${attempt} to load portal...`);

      await page.goto(targetUrl, {
        waitUntil: "domcontentloaded",
        timeout: 60000,
      });
      // ✅ Explicitly wait for login form fields
      await page.waitForSelector("#userID", { timeout: 60000 });
      await page.waitForSelector("#password", { timeout: 60000 });

      console.log("[loginUnify] Page loaded, filling credentials...");
      await page.fill("#userID", username);
      await page.fill("#password", password);
      await page.click("#login");
      // ✅ Post-login check (give portal up to 5 minutes to redirect)
      await page.waitForURL(
        "https://unify.ap.iriworldwide.com/client1/plus/landing/0",
        { timeout: 300_000 }
      );
      console.log("Login successful.");
      return; // success → exit function
    } catch (err) {
      console.error(`[loginUnify] Attempt ${attempt} failed:`, err.message);
      if (attempt >= 2) throw err; // rethrow after max retries
      console.log("[loginUnify] Retrying navigation...");
    }
  }
}

// The master list of target flat files grouped by where to find them in UI
const TARGET_GROUPS = [
  {
    navigate: async (page) => {
      // Favorites → Flat Files 2 → open any thumb to load the dashboard
      await locateAndAction(page, "#FavoritesLink", {
        description: "Favorites",
      });
      await locateAndAction(page, "span", {
        hasText: "Flat Files 2",
        description: "Flat Files 2",
      });
      await locateAndAction(page, "div.thumb-box", {
        hasText: "Flat File - CD",
        description: "Open CD card",
      });
    },
    files: [
      "Flat File - CD",
      "Flat File - NWNI",
      "Flat File - PSNI",
      "Flat File - FSSI",
      "Flat File - Petrol CDNISI",
    ],
  },
  {
    navigate: async (page) => {
      await locateAndAction(page, "#FavoritesLink", {
        description: "Favorites",
      });
      await locateAndAction(page, "span", {
        hasText: "Flat File - TSM NI/SI",
        description: "Flat File - TSM NI/SI",
      });
    },
    files: ["Flat File - TSM NI/SI"],
  },
  {
    navigate: async (page) => {
      await locateAndAction(page, "#FavoritesLink", {
        description: "Favorites",
      });
      await locateAndAction(page, "span", {
        hasText: "Flat File - Chemist Warehouse",
        description: "Flat File - Chemist Warehouse",
      });
    },
    files: ["Flat File - Chemist Warehouse"],
  },
];

async function exportFlatFile(page, fileName) {
  console.log(`\n— Exporting ${fileName} —`);
  // Switch to tab
  await locateAndAction(page, "ul#report-nav-scroll li.reportNavLi", {
    hasText: fileName,
    description: `${fileName} tab`,
  });

  // Action → Export
  await locateAndAction(page, "div.dashboard-action span.db-action-link", {
    hasText: "Action",
    description: "Action button",
  });
  await locateAndAction(
    page,
    "#reportContainer div.actionModal li.action-modal-item span",
    { hasText: "Export", description: "Export option" }
  );

  // SelectAll Geography and Time
  await locateAndAction(page, "div.selectAll label.check-label", {
    action: "check",
    description: "SelectAll (Geo)",
  });
  await locateAndAction(page, "div.iterate-select select", {
    action: "select",
    option: "1: Object",
    description: "Time option",
  });
  await locateAndAction(page, "div.selectAll label.check-label", {
    action: "check",
    description: "SelectAll (Time)",
  });

  // Excel + Pivot Table
  await locateAndAction(page, "div.fileType label.check-label span", {
    hasText: "Excel Spreadsheet",
    action: "check",
    description: "Excel file type",
  });
  await locateAndAction(page, "ul li label.check-label span", {
    hasText: "Pivot Table",
    action: "check",
    description: "Pivot Table",
  });

  // Export & dismiss modal
  await locateAndAction(page, "div.modal-footer div.exp-footer-button button", {
    hasText: "Export",
    description: "Export button",
  });
  await locateAndAction(page, "div.modal-dialog div.modal-content div button", {
    hasText: "Okay",
    description: "Okay button",
  });
}

async function navigateAndQueueExports(page) {
  // Walk groups and trigger exports for each file
  for (const group of TARGET_GROUPS) {
    await group.navigate(page);
    for (const file of group.files) {
      await exportFlatFile(page, file);
    }
  }
}

// ---------- Notifications scraping ----------
const TARGET_FILES = [
  "Flat File - CD",
  "Flat File - NWNI",
  "Flat File - PSNI",
  "Flat File - FSSI",
  "Flat File - Petrol CDNISI",
  "Flat File - TSM NI/SI",
  "Flat File - Chemist Warehouse",
];

async function getNotificationItems(page) {
  await locateAndAction(page, "a.fa-bell", {
    description: "Notifications bell",
  });
  await locateAndAction(page, "a.fa-expand span", {
    hasText: "View All",
    description: "View All",
  });

  await page.waitForSelector("div.cdk-virtual-scroll-content-wrapper div.row", {
    timeout: 10_000,
  });
  const rows = await page
    .locator("div.cdk-virtual-scroll-content-wrapper div.row")
    .all();
  const items = [];
  for (const row of rows) {
    try {
      const titleSpan = row.locator("div.title span.ellipsis");
      const titleAttr = await titleSpan.getAttribute("title");
      const fileName = (titleAttr || (await titleSpan.innerText())).trim();
      if (!TARGET_FILES.includes(fileName)) continue;
      const timeStr = (
        await row.locator("div.col-2").nth(0).innerText()
      ).trim();
      items.push({ locator: titleSpan, fileName, timeStr });
    } catch (e) {
      console.warn("Failed parsing one notification row:", e);
    }
  }
  const latest = items.slice(0, 7);
  console.log("Will download:");
  latest.forEach((i) => console.log(`${i.fileName}, ${i.timeStr}`));
  return latest;
}

// ---------- Parallel downloads with retry ----------
async function clickAndAwaitDownload(
  page,
  locator,
  { timeout = 600_000 } = {}
) {
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout }),
    locator.click(),
  ]);
  return download;
}

async function saveDownload(download, downloadDir, safeName) {
  ensureDir(downloadDir);
  const target = uniquePath(downloadDir, `${safeName}.xlsx`);
  await download.saveAs(target);
  console.log(`✔ Saved ${path.basename(target)}`);
  return target;
}

async function downloadFromNotifications(
  page,
  downloadDir,
  { retries = 2, parallel = true } = {}
) {
  const items = await getNotificationItems(page);
  if (!items.length) {
    console.log("No matching notifications found.");
    return { successes: [], failures: [] };
  }

  const tasks = items.map((item, idx) =>
    (async () => {
      const safeBase = sanitizeFilename(item.fileName);
      for (let attempt = 0; attempt <= retries; attempt++) {
        try {
          const download = await clickAndAwaitDownload(page, item.locator);
          const savedPath = await saveDownload(download, downloadDir, safeBase);
          return { ok: true, file: savedPath, name: item.fileName };
        } catch (err) {
          console.warn(
            `Download failed for ${item.fileName} (attempt ${attempt + 1}/${
              retries + 1
            }):`,
            err?.message || err
          );
          if (attempt === retries)
            return { ok: false, error: String(err), name: item.fileName };
          // Small backoff before retry
          await new Promise((r) => setTimeout(r, 2_000 * (attempt + 1)));
        }
      }
    })()
  );

  // Fire all promises (parallel) and wait
  const results = await Promise.allSettled(tasks);
  const successes = [];
  const failures = [];
  for (const r of results) {
    if (r.status === "fulfilled" && r.value?.ok) successes.push(r.value);
    else failures.push(r.value || r.reason);
  }
  console.log(
    `Downloads complete. Success: ${successes.length}, Failures: ${failures.length}`
  );
  return { successes, failures };
}

// ---------- Move download file to destination ----------
function moveTargetFiles(downloadDir, destinationDir, targetFiles) {
  ensureDir(destinationDir);
  const moved = [];

  for (const baseName of targetFiles) {
    const pattern = `${baseName}.xlsx`;
    const srcPath = path.join(downloadDir, pattern);

    if (fs.existsSync(srcPath)) {
      const destPath = path.join(destinationDir, pattern);

      // Replace old file if exists
      fs.renameSync(srcPath, destPath);
      moved.push({ from: srcPath, to: destPath });

      console.log(`📂 Replaced ${pattern} in destination`);
    } else {
      console.warn(`⚠ Target file not found in downloadDir: ${pattern}`);
    }
  }

  return moved;
}

// ---------- UUID cleanup (today only) ----------
function cleanupUuidFiles(directory) {
  const uuidRe =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.[^.]+$/i;
  const today = new Date().toDateString();
  let removed = 0;
  for (const file of fs.readdirSync(directory)) {
    const full = path.join(directory, file);
    const stat = fs.statSync(full);
    if (!stat.isFile()) continue;
    if (!uuidRe.test(file)) continue;
    const mday = new Date(stat.mtime).toDateString();
    if (mday === today) {
      try {
        fs.unlinkSync(full);
        removed++;
        console.log(`🗑️ Deleted UUID file: ${file}`);
      } catch (e) {
        console.warn(`Couldn't delete ${file}:`, e);
      }
    }
  }
  return removed;
}

// ---------- Utilities ----------
function sanitizeFilename(name) {
  return name.replace(/[\\/*?:"<>|]/g, "_");
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function uniquePath(directory, filename) {
  const ext = path.extname(filename);
  const base = path.basename(filename, ext);
  let attempt = 0;
  let full = path.join(directory, filename);
  while (fs.existsSync(full)) {
    attempt += 1;
    full = path.join(directory, `${base} (${attempt})${ext}`);
  }
  return full;
}

// Helpful locator wrapper
async function locateAndAction(
  page,
  selector,
  {
    hasText,
    action = "click", // click | check | select
    option = undefined,
    description = "",
    timeout = 15_000,
  } = {}
) {
  const loc = hasText
    ? page.locator(selector, { hasText })
    : page.locator(selector);
  await loc.waitFor({ state: "visible", timeout });
  if (action === "click") await loc.click();
  else if (action === "check") await loc.check();
  else if (action === "select") await loc.selectOption(option);
  else throw new Error(`Unsupported action: ${action}`);
  if (description) console.log(`${action} → ${description}`);
  return loc;
}

// ---------- Orchestration ----------
(async () => {
  console.log("[download.js] == script started == ");
  const waitAfterExportMs = 2.5 * 60 * 60 * 1000;
  try {
    // browser creation
    const chromePaths = [
      "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
      "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
      path.join(
        process.env.LOCALAPPDATA || "",
        "Google\\Chrome\\Application\\chrome.exe"
      ),
    ];

    const chromeExecutable = chromePaths.find((p) => fs.existsSync(p));
    if (!chromeExecutable) {
      throw new Error("Chrome executable not found.");
    }

    const browser = await chromium.launch({
      headless: false,
      executablePath: chromeExecutable,
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-accelerated-2d-canvas",
        "--no-first-run",
        "--no-zygote",
      ],
    });

    const page = await browser.newPage();
    // end browser and page created.

    browser.on("disconnected", () => {
      console.error("[Automation ERROR] Browser was closed by the user.");
      process.exit(2); // distinct exit code for user abort
    });

    const downloadDir = process.argv[2];
    const destinationDir = process.argv[3];
    ensureDir(downloadDir);
    ensureDir(destinationDir);

    const username = process.argv[4];
    const password = process.argv[5];

    if (!username || !password) {
      console.error("Missing username/password in configuration");
      process.exit(1);
    }

    // 1) Login
    await loginUnify(page, { username, password });

    // 2) Navigate + queue exports (sequential UI actions)
    await navigateAndQueueExports(page);

    // 3) Wait for server-side exports to complete
    console.log(
      `⏳ Waiting for exports to complete (~${Math.round(
        waitAfterExportMs / 60000
      )} min)...`
    );
    await page.waitForTimeout(waitAfterExportMs);

    // 4) Fetch notifications and download in parallel with retries
    await downloadFromNotifications(page, downloadDir, {
      retries: 2,
      parallel: true,
    });

    // 5) Move real downloads to destination folder
    moveTargetFiles(downloadDir, destinationDir, TARGET_FILES);

    // 6) Cleanup UUID temp files (today only) in downloadDir
    cleanupUuidFiles(downloadDir);

    console.log("Automation finished successfully.");
    await browser.close();
    process.exit(0);
  } catch (err) {
    if (
      err.message.includes("Target page, context or browser has been closed")
    ) {
      console.error("[Automation ERROR] User closed the browser window.");
      process.exit(2);
    } else {
      console.error("[Automation ERROR] Automation failed:", err);
      process.exit(1);
    }
  }
})();
