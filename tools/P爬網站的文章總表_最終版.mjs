import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import path from "path";
import readline from "readline";
import * as XLSX from "xlsx";

// ========= 你可以改的設定 =========
const OUT_DIR = "F:\\F\\AI";                 
const DEFAULT_URL = "https://futuresinvest90223.com/";
const MAX_PAGES = 200;
const SLEEP_MS = 800;
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

// ✅ 關鍵：檔名一律「無空白」
function safeFilename(name, maxLen = 120) {
  name = String(name)
    .replace(/[<>:"/\\|?*]/g, "_")  // 非法字元
    .replace(/\s+/g, "_")           // 空白 → _
    .replace(/_+/g, "_")            // 多個 _ 合併
    .trim();

  if (name.length > maxLen) name = name.slice(0, maxLen);
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
// 1) 取 HTML：axios → Playwright fallback
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
    Referer: `${u.origin}/`,
  };
}

async function fetchHtmlByAxios(url) {
  const res = await axios.get(url, {
    headers: browserLikeHeaders(url),
    timeout: 30000,
    responseType: "text",
    validateStatus: () => true,
  });

  if (res.status !== 200) {
    const err = new Error(`axios status ${res.status}`);
    err.status = res.status;
    throw err;
  }
  return String(res.data || "");
}

async function fetchHtmlByPlaywright(url) {
  const { chromium } = await import("playwright");
  const headless = process.env.HEADLESS === "0" ? false : true;

  const browser = await chromium.launch({ headless });
  const context = await browser.newContext({
    userAgent: browserLikeHeaders(url)["User-Agent"],
    locale: "zh-TW",
    viewport: { width: 1366, height: 900 },
  });

  const page = await context.newPage();
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
  await page.waitForTimeout(1200);

  const html = await page.content();
  await browser.close();
  return html;
}

async function getHtml(url) {
  try {
    return await fetchHtmlByAxios(url);
  } catch {
    console.log(`⚠️ axios 失敗，改用 Playwright：${url}`);
    return await fetchHtmlByPlaywright(url);
  }
}

// ========================
// 2) 解析文章 + 翻頁
// ========================
function extractPosts($, baseUrl) {
  const rows = [];

  $("article").each((_, el) => {
    const a = $(el).find("h1 a, h2 a, .entry-title a").first();
    const title = a.text().trim();
    const link = normalizeUrl(baseUrl, a.attr("href") || "");
    if (!title || !link) return;

    let dt = "";
    const t = $(el).find("time[datetime]").attr("datetime");
    if (t) dt = t.slice(0, 10);

    rows.push({ 日期: dt, 名稱: title, 網址: link });
  });

  const uniq = new Map();
  for (const r of rows) uniq.set(r.網址, r);
  return [...uniq.values()];
}

function findNextPageUrl($, baseUrl) {
  const href =
    $("a.next.page-numbers").attr("href") ||
    $('a[rel="next"]').attr("href");

  return href ? normalizeUrl(baseUrl, href) : "";
}

// ========================
// 3) 輸出 XLSX
// ========================
function writeXlsx(rows, outPath) {
  const ws = XLSX.utils.json_to_sheet(rows, { header: ["日期", "名稱", "網址"] });
  ws["!cols"] = [{ wch: 12 }, { wch: 60 }, { wch: 90 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "posts");
  XLSX.writeFile(wb, outPath);
}

// ========================
// MAIN
// ========================
async function main() {
  let startUrl = await askUrl();
  if (!/^https?:\/\//i.test(startUrl)) startUrl = "https://" + startUrl;

  fs.mkdirSync(OUT_DIR, { recursive: true });

  const seen = new Set();
  const all = [];
  let url = startUrl;

  for (let i = 1; i <= MAX_PAGES; i++) {
    console.log(`🌐 第 ${i} 頁：${url}`);

    const html = await getHtml(url);
    const $ = cheerio.load(html);
    const rows = extractPosts($, url);

    for (const r of rows) {
      if (!seen.has(r.網址)) {
        seen.add(r.網址);
        all.push(r);
      }
    }

    const next = findNextPageUrl($, url);
    if (!next) break;
    url = next;
    await sleep(SLEEP_MS);
  }

  const outPath = path.join(OUT_DIR, filenameFromUrl(startUrl));
  writeXlsx(all, outPath);

  console.log(`🎉 完成，共 ${all.length} 筆`);
  console.log(`📄 檔案：${outPath}`);
}

main().catch(console.error);
