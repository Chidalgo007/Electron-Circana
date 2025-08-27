const puppeteer = require("puppeteer");

(async () => {
  try {
    const browser = await puppeteer.launch({
      headless: false,
      executablePath:
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", // adjust path
    });

    const page = await browser.newPage();
    await page.goto("https://example.com");

    console.log("Automation started: visiting website...");

    // Example click
    await page.waitForSelector("#download");
    await page.click("#download");

    console.log("Automation finished successfully.");
    await browser.close();
    process.exit(0);
  } catch (err) {
    console.error("Automation failed:", err);
    process.exit(1);
  }
})();
