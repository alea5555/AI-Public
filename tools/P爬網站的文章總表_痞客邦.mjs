import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import path from "path";
import readline from "readline";
import * as XLSX from "xlsx";

const OUT_DIR = "F:\\F\\AI";
const PAGE_SLEEP_MS = 800;        // 翻頁間隔
const ARTICLE_SLEEP_MS = 250;     // 補抓日期間隔（避免太快）
const STOP_EMPTY_PAGES = 3;       // 連續幾頁沒有新文章就停止
const MAX_PAGES_GUARD = 1000;     // 保險上限，避免無限迴圈

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function askUrl() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question("請輸入要爬取的 PIXNET 網址（首頁或分類頁都可）：\n", (answer) => {
      rl.close();
      resolve((answer || "").trim());
    });
  });
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

function setPage(url, page) {
  const u = new URL(url);
  u.searchParams.set("page", String(page));
  return u.toString();
}

// 只做「列表頁」抓文章：標題、連結、（若剛好有）日期
function extractPostsFromList($) {
  const rows = [];

  // 以 main/container 優先，避免抓到側欄熱門文章/分類列表
  const root = $("main").length ? $("main") : $("#container").length ? $("#container") : $.root();

  let articles = root.find("article");
  if (articles.length === 0) articles = root.find("h2");

  articles.each((_, el) => {
    const $el = $(el);

    const a = $el.is("h2") ? $el.find("a").first() : $el.find("h2 a").first();
    const link = (a.attr("href") || "").trim();
    const title = a.text().trim();
    if (!link || !title) return;

    // 列表頁日期（若版型沒有就會空，後面會補抓）
    let dt = "";
    const timeEl = $el.find("time[datetime]").first();
    if (timeEl.length) dt = (timeEl.attr("datetime") || "").slice(0, 10);

    rows.push({ 日期: dt, 名稱: title, 網址: link });
  });

  // 以網址去重
  const uniq = new Map();
  for (const r of rows) uniq.set(r.網址, r);
  return [...uniq.values()];
}

// ---------- 補抓日期（進文章頁） ----------
function parseDateToYYYYMMDD(s) {
  if (!s) return "";
  const str = String(s).trim();

  let m = str.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const y = m[1];
    const mo = String(parseInt(m[2], 10)).padStart(2, "0");
    const d = String(parseInt(m[3], 10)).padStart(2, "0");
    return `${y}-${mo}-${d}`;
  }

  m = str.match(/(20\d{2})\/(\d{1,2})\/(\d{1,2})/);
  if (m) {
    const y = m[1];
    const mo = String(parseInt(m[2], 10)).padStart(2, "0");
    const d = String(parseInt(m[3], 10)).padStart(2, "0");
    return `${y}-${mo}-${d}`;
  }

  return "";
}

function extractDateFromArticleHtml(html) {
  const $ = cheerio.load(html);

  // 常見 meta
  const metaChecks = [
    ["property", "article:published_time"],
    ["property", "og:published_time"],
    ["name", "pubdate"],
    ["name", "publishdate"],
    ["name", "publish_date"],
    ["name", "date"],
    ["itemprop", "datePublished"],
  ];

  for (const [attr, val] of metaChecks) {
    const tag = $(`meta[${attr}="${val}"]`).first();
    const d = parseDateToYYYYMMDD(tag.attr("content"));
    if (d) return d;
  }

  // time
  const t = $("time").first();
  if (t.length) {
    const d = parseDateToYYYYMMDD(t.attr("datetime") || t.text());
    if (d) return d;
  }

  // regex fallback
  const m = html.match(/\b(20\d{2})[-/](\d{1,2})[-/](\d{1,2})\b/);
  if (m) {
    const y = m[1];
    const mo = String(parseInt(m[2], 10)).padStart(2, "0");
    const d = String(parseInt(m[3], 10)).padStart(2, "0");
    return `${y}-${mo}-${d}`;
  }

  return "";
}

async function fetchArticleDate(axiosInst, url) {
  try {
    const html = (await axiosInst.get(url)).data;
    return extractDateFromArticleHtml(html);
  } catch {
    return "";
  }
}

// ---------- 輸出 xlsx ----------
function writeXlsx(rows, outPath) {
  const data = rows.map((r) => ({
    日期: r.日期 || "",
    名稱: r.名稱 || "",
    網址: r.網址 || "",
  }));

  const ws = XLSX.utils.json_to_sheet(data, { header: ["日期", "名稱", "網址"] });
  ws["!cols"] = [{ wch: 12 }, { wch: 40 }, { wch: 70 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "posts");
  XLSX.writeFile(wb, outPath);
}

async function main() {
  let startUrl = process.argv[2];
  if (!startUrl) startUrl = await askUrl();

  if (!startUrl) {
    console.log("❌ 未輸入網址，結束程式");
    process.exit(1);
  }
  if (!/^https?:\/\//i.test(startUrl)) startUrl = "https://" + startUrl;

  fs.mkdirSync(OUT_DIR, { recursive: true });

  const axiosInst = axios.create({
    headers: { "User-Agent": "Mozilla/5.0" },
    timeout: 30000,
  });

  // ✅ 不依賴 maxPage：一直翻到沒新文章為止
  const seen = new Set();
  const all = [];
  let emptyPages = 0;

  for (let page = 1; page <= MAX_PAGES_GUARD; page++) {
    const pageUrl = setPage(startUrl, page);

    let html = "";
    try {
      html = (await axiosInst.get(pageUrl)).data;
    } catch (e) {
      // 取不到頁面就視為結尾
      console.log(`⚠️ page=${page} 讀取失敗，停止翻頁`);
      break;
    }

    const $ = cheerio.load(html);
    const rows = extractPostsFromList($);

    let added = 0;
    for (const r of rows) {
      if (!seen.has(r.網址)) {
        seen.add(r.網址);
        all.push(r);
        added++;
      }
    }

    console.log(`page=${page} 本頁新增 ${added} 累計 ${all.length}`);

    if (added === 0) emptyPages++;
    else emptyPages = 0;

    if (emptyPages >= STOP_EMPTY_PAGES) {
      console.log(`📌 連續 ${STOP_EMPTY_PAGES} 頁無新文章，停止翻頁`);
      break;
    }

    await sleep(PAGE_SLEEP_MS);
  }

  console.log(`✅ 文章列表蒐集完成：${all.length} 篇`);

  // 補抓日期
  let needFix = all.filter((r) => !r.日期).length;
  console.log(`需要補抓日期：${needFix} 篇（會逐篇進文章頁抓 meta/time）`);

  for (let i = 0; i < all.length; i++) {
    if (all[i].日期) continue;

    const d = await fetchArticleDate(axiosInst, all[i].網址);
    if (d) all[i].日期 = d;

    await sleep(ARTICLE_SLEEP_MS);

    if ((i + 1) % 25 === 0) {
      const left = all.filter((r) => !r.日期).length;
      console.log(`補抓日期進度：${i + 1}/${all.length}，尚未補到日期：${left} 篇`);
    }
  }

  const outFile = filenameFromUrl(startUrl);
  const outPath = path.join(OUT_DIR, outFile);

  writeXlsx(all, outPath);

  console.log(`✅ 完成，共 ${all.length} 筆`);
  console.log(`📄 輸出位置：${outPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
