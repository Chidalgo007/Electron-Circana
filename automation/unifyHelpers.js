// unifyHelpers.js
const fs = require("fs");
const path = require("path");

// ---------- Configuration ----------
const TARGET_FILES = [
  "Flat File - CD",
  "Flat File - NWNI",
  "Flat File - PSNI",
  "Flat File - FSSI",
  "Flat File - Petrol CDNISI",
  "Flat File - TSM NI/SI",
  "Flat File - Chemist Warehouse",
];

// ---------- Login ----------
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
      await page.waitForSelector("#userID", { timeout: 60000 });
      await page.waitForSelector("#password", { timeout: 60000 });

      console.log("[loginUnify] Page loaded, filling credentials...");
      await page.fill("#userID", username);
      await page.fill("#password", password);
      await page.click("#login");
      await page.waitForURL(
        "https://unify.ap.iriworldwide.com/client1/plus/landing/0",
        { timeout: 300_000 }
      );
      console.log("Login successful.");
      return;
    } catch (err) {
      console.error(`[loginUnify] Attempt ${attempt} failed:`, err.message);
      if (attempt >= 2) throw err;
      console.log("[loginUnify] Retrying navigation...");
    }
  }
}

// ---------- Navigation + Export Queue ----------
async function navigateAndQueueExports(page) {
  const TARGET_GROUPS = [
    {
      navigate: async (page) => {
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
    await locateAndAction(page, "ul#report-nav-scroll li.reportNavLi", {
      hasText: fileName,
      description: `${fileName} tab`,
    });
    await locateAndAction(page, "div.dashboard-action span.db-action-link", {
      hasText: "Action",
      description: "Action button",
    });
    await locateAndAction(
      page,
      "#reportContainer div.actionModal li.action-modal-item span",
      { hasText: "Export", description: "Export option" }
    );

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

    await locateAndAction(
      page,
      "div.modal-footer div.exp-footer-button button",
      { hasText: "Export", description: "Export button" }
    );
    await locateAndAction(
      page,
      "div.modal-dialog div.modal-content div button",
      { hasText: "Okay", description: "Okay button" }
    );
  }

  for (const group of TARGET_GROUPS) {
    await group.navigate(page);
    for (const file of group.files) {
      await exportFlatFile(page, file);
    }
  }
}

// ---------- Notifications / Download ----------
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

  const tasks = items.map((item) =>
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
          await new Promise((r) => setTimeout(r, 2_000 * (attempt + 1)));
        }
      }
    })()
  );

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

// ---------- File management ----------
function moveTargetFiles(downloadDir, destinationDir, targetFiles) {
  ensureDir(destinationDir);
  const moved = [];
  for (const baseName of targetFiles) {
    const srcPath = path.join(downloadDir, `${baseName}.xlsx`);
    if (fs.existsSync(srcPath)) {
      const destPath = path.join(destinationDir, `${baseName}.xlsx`);
      fs.renameSync(srcPath, destPath);
      moved.push({ from: srcPath, to: destPath });
      console.log(`📂 Replaced ${baseName}.xlsx in destination`);
    } else
      console.warn(`⚠ Target file not found in downloadDir: ${baseName}.xlsx`);
  }
  return moved;
}

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
    if (new Date(stat.mtime).toDateString() === today) {
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
  const ext = path.extname(filename),
    base = path.basename(filename, ext);
  let attempt = 0,
    full = path.join(directory, filename);
  while (fs.existsSync(full)) {
    attempt++;
    full = path.join(directory, `${base} (${attempt})${ext}`);
  }
  return full;
}

async function locateAndAction(
  page,
  selector,
  { hasText, action = "click", option, description = "", timeout = 15_000 } = {}
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

// ---------- Exports ----------
module.exports = {
  loginUnify,
  navigateAndQueueExports,
  getNotificationItems,
  downloadFromNotifications,
  moveTargetFiles,
  cleanupUuidFiles,
  ensureDir,
  sanitizeFilename,
  uniquePath,
  locateAndAction,
  TARGET_FILES,
};
