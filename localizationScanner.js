// ============================================================
// Localization Scanner — Playwright JavaScript Version
// Converted from Java Selenium by Claude
//
// WHAT THIS SCRIPT DOES:
//   1. Opens an Arabic-language LMS website
//   2. Logs in and dismisses popups
//   3. Collects all internal links on the homepage
//   4. Visits each link and checks if English text is present
//      (since the site should be fully in Arabic)
//   5. Saves the findings in an Excel (.xlsx) file
// ============================================================

const { chromium } = require("playwright");  // Playwright browser automation
const ExcelJS = require("exceljs");           // Library to write Excel files
//git testing
// ── Configuration ────────────────────────────────────────────
const BASE_URL = "https://altasnim.uknowva-stage.in/?lang=ar-AA";
const MAX_URLS = 10; // Limit scan to first 10 URLs (for testing)

// This array will collect all findings throughout the scan
const reportData = []; // Each entry: { url, findings }

// ── Main Entry Point ─────────────────────────────────────────
(async () => {
  // Launch a visible Chrome browser window
  // In Java Selenium: new ChromeDriver(options)
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    await page.setViewportSize({ width: 1280, height: 800 });
    await page.goto(BASE_URL);

    await performLogin(page);
    await handleInitialPopups(page);

    const uniqueUrls = await collectLinks(page);

    // Take only the first MAX_URLS from the collected links
    // In Java: uniqueUrls.stream().limit(MAX_URLS).collect(...)
    const limitedUrls = [...uniqueUrls].slice(0, MAX_URLS);

    console.log(`Total unique URLs found: ${uniqueUrls.size}`);
    console.log(`Scanning first ${limitedUrls.length} URLs...\n`);

    let count = 1;
    for (const url of limitedUrls) {
      console.log(`[${count++}/${limitedUrls.length}] Scanning: ${url}`);
      try {
        await scanUrlForEnglish(page, url);
      } catch (err) {
        // If any page fails, record the error instead of crashing
        reportData.push({ url, findings: `Error: ${err.message}` });
      }
    }

    await writeToExcel();
  } finally {
    // Always close the browser, even if an error occurred
    // In Java: driver.quit()
    await browser.close();
  }
})();

// ── Step 1: Login ─────────────────────────────────────────────
// In Java: driver.findElement(By.id("username")).sendKeys("uknowva")
// In Playwright: page.fill("#username", "uknowva")
async function performLogin(page) {
  await page.waitForSelector("#username", { timeout: 30000 });
  await page.fill("#username", "uknowva");
  await page.fill("#password", "9jcmVhl$#@");
  await page.click("button[type='submit']");

  // Wait for URL to change after login (more reliable than waitForLoadState)
  try {
    await page.waitForURL("**", { timeout: 15000 });
  } catch {
    // If URL doesn't change, just wait a few seconds and continue
  }
  await page.waitForTimeout(3000);
}

// ── Step 2: Dismiss any popups after login ────────────────────
async function handleInitialPopups(page) {
  try {
    // Wait up to 5 seconds for a "Skip" link to appear
    // In Java: wait.until(ExpectedConditions.elementToBeClickable(...))
    await page.waitForSelector("a:text('Skip')", { timeout: 5000 });
    await page.click("a:text('Skip')");

    // Handle browser alert if it pops up
    await handleAlertIfPresent(page);
  } catch {
    // If no Skip button found within 5s, just continue
    console.log("No initial Skip button found.");
  }
}

// ── Step 3: Collect all links on the page ────────────────────
// In Java: driver.findElements(By.tagName("a"))
// In Playwright: page.$$eval("a", ...)
async function collectLinks(page) {
  const hrefs = await page.$$eval("a", (anchors) =>
    anchors
      .map((a) => a.href)
      // Keep only absolute http(s) links, exclude logout
      .filter((href) => href && href.startsWith("http") && !href.includes("logout"))
  );

  // Use a Set to automatically remove duplicates
  // In Java: Collectors.toSet()
  return new Set(hrefs);
}

// ── Step 4: Visit a URL and check for English text ───────────
async function scanUrlForEnglish(page, url) {
  await page.goto(url);

  // Wait 3 seconds for dynamic content to load
  // In Java: Thread.sleep(3000)
  await page.waitForTimeout(3000);

  // Handle any alert dialog that might appear
  await handleAlertIfPresent(page);

  // Wait for the page body to exist before reading text
  // In Java: wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("body")))
  await page.waitForSelector("body");

  // Get all visible text on the page
  // In Java: driver.findElement(By.tagName("body")).getText()
  const pageText = await page.innerText("body");

  // Split text into lines and check each one
  const lines = pageText.split(/\r?\n/);
  let foundAny = false;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue; // Skip empty lines

    if (isEnglishText(trimmed)) {
      reportData.push({ url, findings: trimmed });
      foundAny = true;
    }
  }

  // If no English text found on this page, mark it as clean
  if (!foundAny) {
    reportData.push({ url, findings: "Clean (Arabic Only)" });
  }

  // Count how many English lines were found for this URL
  const englishCount = reportData.filter(
    (r) => r.url === url && !r.findings.startsWith("Clean") && !r.findings.startsWith("Error")
  ).length;

  console.log(`   Done. English lines found: ${englishCount}`);
}

// ── Helper: Detect English text ───────────────────────────────
// Returns true if the line is primarily English/Latin text.
// Rules:
//   - Must have at least 2 Latin characters
//   - Must have at least 2 English words (2+ letters each)
//   - Latin character count >= Arabic character count
//
// In Java: used Pattern/Matcher and stream().filter()
// In JS: we use regex and string iteration
function isEnglishText(text) {
  // Count Latin (English) characters
  const latinChars = [...text].filter(
    (c) => (c >= "A" && c <= "Z") || (c >= "a" && c <= "z")
  ).length;

  // Count Arabic characters (Unicode range 0x0600–0x06FF)
  const arabicChars = [...text].filter((c) => {
    const code = c.charCodeAt(0);
    return code >= 0x0600 && code <= 0x06ff;
  }).length;

  if (latinChars < 2) return false;

  // Count English words (sequences of 2+ letters)
  // In Java: Pattern.compile("[a-zA-Z]{2,}")
  const wordMatches = text.match(/[a-zA-Z]{2,}/g) || [];
  const wordCount = wordMatches.length;

  return wordCount >= 2 && latinChars >= arabicChars;
}

// ── Helper: Dismiss browser alerts if present ────────────────
// In Java: driver.switchTo().alert().accept()
// In Playwright: page.on("dialog", ...) works for pre-existing dialogs,
// but for one-time handling we use a one-shot listener
async function handleAlertIfPresent(page) {
  // Set up a one-time listener for a dialog (alert/confirm/prompt)
  page.once("dialog", async (dialog) => {
    await dialog.accept(); // Accept/dismiss the alert
  });

  // Small wait to allow any pending alert to fire
  await page.waitForTimeout(500);
}

// ── Step 5: Write results to Excel ───────────────────────────
// In Java: Apache POI (XSSFWorkbook, Sheet, Row, Cell, CellStyle)
// In JS:   ExcelJS (Workbook, Worksheet, Row, Cell)
async function writeToExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Localization Report");

  // ── Define columns with widths ────────────────────────────
  sheet.columns = [
    { header: "Sr. No.",            key: "srNo",     width: 10 },
    { header: "Page URL",           key: "url",      width: 50 },
    { header: "English Text Found", key: "findings", width: 80 },
  ];

  // ── Style the header row (Row 1) ──────────────────────────
  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, size: 11 };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFA500" }, // Orange background
    };
    cell.alignment = { horizontal: "center" };
    cell.border = { bottom: { style: "thin" } };
  });

  // ── Add data rows ─────────────────────────────────────────
  let srNo = 1;
  let lastUrl = "";

  for (const res of reportData) {
    const isNewUrl = res.url !== lastUrl;

    const row = sheet.addRow({
      srNo: srNo++,
      url: isNewUrl ? res.url : "", // Show URL only once per group
      findings: res.findings,
    });

    // Style the URL cell differently when it's the first row for that URL
    const urlCell = row.getCell("url");
    if (isNewUrl) {
      urlCell.font = { bold: true };
      urlCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD3D3D3" }, // Light grey background
      };
      lastUrl = res.url;
    }

    // Wrap text in URL and findings cells
    urlCell.alignment = { wrapText: true, vertical: "top" };
    row.getCell("findings").alignment = { wrapText: true, vertical: "top" };
  }

  // ── Save to file ──────────────────────────────────────────
  const outputPath = "localization_report.xlsx";
  await workbook.xlsx.writeFile(outputPath);

  console.log(`\nExcel file generated: ${outputPath}`);
  console.log(`Total rows written: ${reportData.length}`);
}
