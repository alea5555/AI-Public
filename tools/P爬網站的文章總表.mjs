import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import path from "path";
import readline from "readline";
import * as XLSX from "xlsx";

const SLEEP_MS = 800;
const OUT_DIR = "F:\\F\\AI";
const HEADERS = { "User-Agent": "Mozilla/5.0" };

/* ---------- 工具 ---------- */

const sleep = ms => new Promise(r => setTimeout(r, ms));

function askUrl() {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question("請輸入要爬取的網站網址：\n", ans => {
      rl.close();
      resolve((ans || "").trim());
    });
  });
}

function safeFilename(s) {
  return s.replace(/[<>:"/\\|?*]/g, "_").slice(0, 120);
}

function filenameFromUrl(url) {
  const u = new URL(url);
  return safeFilename(u.hostname + "_articles.xlsx");
}

function isSameDomain(url, base) {
  try {
    return new URL(url).hostname === new URL(base).hostname;
  } catch {
    return false;
  }
}

/* ---------- 抓列表 ---------- */

function extractList($, baseUrl) {
  const rows = [];

  const hasArticle = $("article").length > 0;

  if (hasArticle) {
    // ✅ Pixnet 類
    $("article h2 a").each((_, a) => {
      const title = $(a).text().trim();
      const href = $(a).attr("href");
      if (!title || !href) return;
      rows.push({ 名稱: title, 網址: new URL(href, baseUrl).href });
    });
  } else {
    // ✅ 官網 / 新聞站
    $("a").each((_, a) => {
      const title = $(a).text().trim();
      const href = $(a).attr("href");
      if (!title || title.length < 6 || !href) return;

      let url;
      try {
        url = new URL(href, baseUrl).href;
      } catch {
        return;
      }

      if (!isSameDomain(url, baseUrl)) return;
      if (/javascript:|mailto:|#/.test(url)) return;

      rows.push({ 名稱: title, 網址: url });
    });
  }

  const uniq = new Map();
  rows.forEach(r => uniq.set(r.網址, r));
  return [...uniq.values()];
}

/* ---------- 抓日期 ---------- */

async function fetchDate(url) {
  try {
    const html = (await axios.get(url, { headers: HEADERS })).data;
    const $ = cheerio.load(html);

    const t1 = $("time[datetime]").attr("datetime");
    if (t1) return t1.slice(0, 10);

    const t2 = $('meta[property="article:published_time"]').attr("content");
    if (t2) return t2.slice(0, 10);

    const t3 = $('meta[name*="date"]').attr("content");
    if (t3) return t3.slice(0, 10);

  } catch {}
  return "";
}

/* ---------- Excel ---------- */

function writeXlsx(rows, outPath) {
  const ws = XLSX.utils.json_to_sheet(rows, { header: ["日期", "名稱", "網址"] });
  ws["!cols"] = [{ wch: 12 }, { wch: 40 }, { wch: 70 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "articles");
  XLSX.writeFile(wb, outPath);
}

/* ---------- 主流程 ---------- */

async function main() {
  let startUrl = process.argv[2];
  if (!startUrl) startUrl = await askUrl();
  if (!/^https?:\/\//i.test(startUrl)) startUrl = "https://" + startUrl;

  fs.mkdirSync(OUT_DIR, { recursive: true });

  console.log("🔍 讀取首頁中…");
  const html = (await axios.get(startUrl, { headers: HEADERS })).data;
  const $ = cheerio.load(html);

  const list = extractList($, startUrl);
  console.log(`📄 偵測文章數：${list.length}`);

  let i = 0;
  for (const r of list) {
    i++;
    r.日期 = await fetchDate(r.網址);
    console.log(`(${i}/${list.length}) ${r.名稱} ${r.日期 || "(無日期)"}`);
    await sleep(SLEEP_MS);
  }

  const outPath = path.join(OUT_DIR, filenameFromUrl(startUrl));
  writeXlsx(list, outPath);

  console.log(`\n✅ 完成，共 ${list.length} 筆`);
  console.log(`📄 輸出：${outPath}`);
}

main();
