import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import path from "path";
import readline from "readline";
import * as XLSX from "xlsx";

// ========= 你可以改的設定 =========
const OUT_DIR = "F:\\F\\AI";                 // 輸出目錄
const DEFAULT_URL = "https://futuresinvest90223.com/"; // 你按 Enter 時用的預設網址
const MAX_PAGES = 200;                       // 安全上限，避免無限翻頁
const SLEEP_MS = 800;                        // 翻頁間隔
// =================================

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function ask(question) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(question, (ans) => {
      rl.close();
      resolve((ans || "").trim());
    });
  });
}

async function askUrl() {
  const input = await ask(`請輸入要爬取的網站網址（Enter=預設 ${DEFAULT_URL}）：\n`);
  return input || DEFAULT_URL;
}

function safeFilename(name, maxLen = 120) {
  name = String(name).replace(/[<>:"/\\|?*]/g, "_").trim();
  name = name.replace(/\s+/g, " ");
  if (name.length > maxLen) name = name.slice(0, maxLen).trim();
  return name || "output";
}

function filenameFromUrl(startUrl) {
  const u = new URL(startUrl);
  const host = u.hostname.replace(/\./g, "_");
  const parts = u.pathname.split("/").filter(Boolean);
  const tail = parts.length ? parts.join("_") : "root";
  return safeFilename(`${host}_${tail}`) + ".xlsx";
}

function normalizeUrl(base, href) {
  try {
    return new URL(href, base).toString();
  } catch {
    return "";
  }
}

// ========================
// 1) 取 HTML：axios -> 失敗就 Playwright
// ========================
function browserLikeHeaders(url) {
  const u = new URL(url);
  return {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    Accept:
      "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    Connection: "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    Referer: `${u.origin}/`,
  };
}

async function fetchHtmlByAxios(url) {
  const res = await axios.get(url, {
    headers: browserLikeHeaders(url),
    timeout: 30000,
    responseType: "text",
    decompress: true,
    validateStatus: () => true,
  });

  if (res.status !== 200) {
    const err = new Error(`axios status ${res.status}`);
    err.status = res.status;
    err.data = res.data;
    throw err;
  }

  return String(res.data || "");
}

async function fetchHtmlByPlaywright(url) {
  // 動態載入：只有在需要 fallback 時才要求你裝 playwright
  const { chromium } = await import("playwright");

  // HEADLESS=0 可改成有頭模式（某些站對 headless 更嚴）
  const headless = process.env.HEADLESS === "0" ? false : true;

  const browser = await chromium.launch({
    headless,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-blink-features=AutomationControlled",
    ],
  });

  const context = await browser.newContext({
    userAgent: browserLikeHeaders(url)["User-Agent"],
    locale: "zh-TW",
    extraHTTPHeaders: { "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8" },
    viewport: { width: 1366, height: 900 },
  });

  const page = await context.newPage();
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

  // WordPress 有時候會晚點才補文字
  await page.waitForTimeout(1200);

  const html = await page.content();
  await browser.close();
  return html;
}

async function getHtml(url) {
  try {
    return await fetchHtmlByAxios(url);
  } catch (e) {
    const st = e?.status || 0;

    // 常見：403 / 406 / 429 被擋 → 真瀏覽器
    if ([403, 406, 429].includes(st) || /status\s+(403|406|429)/i.test(String(e?.message))) {
      console.log(`⚠️ axios 被擋（${st || "unknown"}），改用 Playwright：${url}`);
      return await fetchHtmlByPlaywright(url);
    }

    console.log(`⚠️ axios 失敗，改用 Playwright：${url}`);
    return await fetchHtmlByPlaywright(url);
  }
}

// ========================
// 2) 解析文章列表 + 下一頁
// ========================
function extractPosts($, baseUrl) {
  const rows = [];

  // WordPress 常見：article / .entry-title
  const articles = $("article");
  if (articles.length) {
    articles.each((_, el) => {
      const a = $(el).find("h1 a, h2 a, .entry-title a").first();
      const title = a.text().trim();
      const link = normalizeUrl(baseUrl, (a.attr("href") || "").trim());
      if (!title || !link) return;

      // 日期：time[datetime] 或 .entry-date
      let dt = "";
      const timeEl = $(el).find("time[datetime]").first();
      if (timeEl.length) dt = (timeEl.attr("datetime") || "").slice(0, 10);

      if (!dt) {
        const t2 = $(el).find(".entry-date, .posted-on time").first();
        if (t2.length) dt = (t2.attr("datetime") || t2.text() || "").trim().slice(0, 10);
      }

      rows.push({ 日期: dt, 名稱: title, 網址: link });
    });
  } else {
    // 退路：抓同網域且看起來像文章的連結
    const baseHost = new URL(baseUrl).hostname;

    $("a[href]").each((_, el) => {
      const a = $(el);
      const title = a.text().trim();
      const href = (a.attr("href") || "").trim();
      const link = normalizeUrl(baseUrl, href);
      if (!title || !link) return;

      try {
        const u = new URL(link);
        if (u.hostname !== baseHost) return;
        if (u.pathname === "/") return;
        if (u.pathname.startsWith("/category/") || u.pathname.startsWith("/tag/")) return;
      } catch {
        return;
      }

      rows.push({ 日期: "", 名稱: title, 網址: link });
    });
  }

  // 同網址去重
  const uniq = new Map();
  for (const r of rows) uniq.set(r.網址, r);
  return [...uniq.values()];
}

function findNextPageUrl($, baseUrl) {
  // WordPress 常見
  let href =
    ($("a.next.page-numbers").attr("href") || "").trim() ||
    ($('a[rel="next"]').attr("href") || "").trim();

  if (href) return normalizeUrl(baseUrl, href);

  // 再退路：pagination 裡面找「下一頁」字樣
  const candidates = [];
  $("a[href]").each((_, el) => {
    const t = $(el).text().trim();
    const h = ($(el).attr("href") || "").trim();
    if (!h) return;
    if (/(下一頁|Next|›|»)/i.test(t)) candidates.push(h);
  });

  if (candidates.length) return normalizeUrl(baseUrl, candidates[0]);
  return "";
}

// ========================
// 3) 輸出 XLSX
// ========================
function writeXlsx(rows, outPath) {
  const data = rows.map((r) => ({
    日期: r.日期 || "",
    名稱: r.名稱 || "",
    網址: r.網址 || "",
  }));

  const ws = XLSX.utils.json_to_sheet(data, { header: ["日期", "名稱", "網址"] });
  ws["!cols"] = [{ wch: 12 }, { wch: 60 }, { wch: 90 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "posts");
  XLSX.writeFile(wb, outPath);
}

// ========================
// MAIN
// ========================
async function main() {
  let startUrl = await askUrl(); // ✅ 一律由你輸入，不吃 argv，不在指令列寫死網址

  if (!startUrl) {
    console.log("❌ 未輸入網址，結束程式");
    process.exit(1);
  }

  if (!/^https?:\/\//i.test(startUrl)) startUrl = "https://" + startUrl;

  fs.mkdirSync(OUT_DIR, { recursive: true });

  const seen = new Set();
  const all = [];

  let url = startUrl;

  for (let pageNo = 1; pageNo <= MAX_PAGES; pageNo++) {
    console.log(`\n🌐 讀取第 ${pageNo} 頁：${url}`);

    const html = await getHtml(url);
    const $ = cheerio.load(html);

    const rows = extractPosts($, url);

    let add = 0;
    for (const r of rows) {
      if (!seen.has(r.網址)) {
        seen.add(r.網址);
        all.push(r);
        add++;
      }
    }

    console.log(`✅ 本頁抓到 ${rows.length} 筆，新增 ${add} 筆，累計 ${all.length} 筆`);

    const nextUrl = findNextPageUrl($, url);
    if (!nextUrl || nextUrl === url) {
      console.log("⛳ 找不到下一頁，結束翻頁。");
      break;
    }

    url = nextUrl;
    await sleep(SLEEP_MS);
  }

  const outFile = filenameFromUrl(startUrl);
  const outPath = path.join(OUT_DIR, outFile);
  writeXlsx(all, outPath);

  console.log(`\n🎉 完成，共 ${all.length} 筆`);
  console.log(`📄 輸出位置：${outPath}`);

  console.log(`\nℹ️ 若你遇到站方擋 headless，可用「有頭模式」執行：`);
  console.log(`   set HEADLESS=0`);
  console.log(`   node ${path.basename(process.argv[1])}`);
}

main().catch((e) => {
  console.error("❌ 程式失敗：", e);
  process.exit(1);
});
