// 檔名：get_pixnet_date_playwright.mjs
import { chromium } from "playwright";

const MONTH = {
  jan:"01", feb:"02", mar:"03", apr:"04", may:"05", jun:"06",
  jul:"07", aug:"08", sep:"09", oct:"10", nov:"11", dec:"12"
};

async function main() {
  const url = process.argv[2];
  if (!url) {
    console.log("用法：node get_pixnet_date_playwright.mjs <url>");
    process.exit(1);
  }

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

  // 等 publish 出現（最多等 10 秒）
  await page.waitForSelector("li.publish span.year", { timeout: 10000 }).catch(() => {});

  const year = (await page.locator("li.publish span.year").first().textContent() || "").trim();
  const monRaw = (await page.locator("li.publish span.month").first().textContent() || "").trim().toLowerCase();
  const dayRaw = (await page.locator("li.publish span.date").first().textContent() || "").trim();

  const mon = MONTH[monRaw] || "";
  const day = dayRaw.replace(/\D/g, "").padStart(2, "0");

  const date8 = (year && mon && day) ? `${year}${mon}${day}` : "";
  console.log(date8 || "抓不到日期");

  await browser.close();
}

main();
