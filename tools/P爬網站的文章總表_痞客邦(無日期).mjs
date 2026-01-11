// 檔名：P爬網站的文章總表_痞客邦.mjs
import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import path from "path";
import readline from "readline";
import * as XLSX from "xlsx";

const SLEEP_MS = 800;
const OUT_DIR = "F:\\F\\AI"; // ✅ 固定輸出到 F:\F\AI

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function askUrl() {
  
  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    rl.question("請輸入要爬取的 PIXNET 網址：\n", (answer) => {
      rl.close();
      resolve((answer || "").trim());
    });
  });
}

function safeFilename(name, maxLen = 120) {
  // Windows 不可用字元 <>:"/\|?*
  name = String(name).replace(/[<>:"/\\|?*]/g, "_").trim();
  name = name.replace(/\s+/g, " ");
  if (name.length > maxLen) name = name.slice(0, maxLen).trim();
  return name || "output";
}

function filenameFromUrl(startUrl) {
  // 例：https://ey90223.pixnet.net/blog  → ey90223_pixnet_net_blog.xlsx
  const u = new URL(startUrl);
  const host = u.hostname.replace(/\./g, "_");
  const parts = u.pathname.split("/").filter(Boolean); // ["blog"] 或更多
  const tail = parts.length ? parts.join("_") : "root";
  return safeFilename(`${host}_${tail}`) + ".xlsx";
}

function setPage(url, page) {
  const u = new URL(url);
  u.searchParams.set("page", String(page));
  return u.toString();
}

function getMaxPage($) {
  let maxPage = 1;

  $("a[href]").each((_, el) => {
    const href = $(el).attr("href") || "";

    // 1) 從 href 的 page= 解析
    try {
      const u = new URL(href, "https://dummy.base/");
      const p = u.searchParams.get("page");
      if (p && /^\d+$/.test(p)) maxPage = Math.max(maxPage, parseInt(p, 10));
    } catch {}

    // 2) 從頁碼文字解析（例如 "43"）
    const txt = $(el).text().trim();
    if (/^\d+$/.test(txt)) maxPage = Math.max(maxPage, parseInt(txt, 10));
  });

  return maxPage;
}

function extractPosts($) {
  const rows = [];

  // 常見版型是 article，也有可能只剩 h2
  let articles = $("article");
  if (articles.length === 0) articles = $("h2");

  articles.each((_, el) => {
    const a = $(el).is("h2")
      ? $(el).find("a").first()
      : $(el).find("h2 a").first();

    const link = (a.attr("href") || "").trim();
    const title = a.text().trim();
    if (!link || !title) return;

    // 日期：優先抓 time[datetime]
    let dt = "";
    const timeEl = $(el).find("time[datetime]").first();
    if (timeEl.length) dt = (timeEl.attr("datetime") || "").slice(0, 10);

    rows.push({ 日期: dt, 名稱: title, 網址: link });
  });

  // 同網址去重（保留最後一次）
  const uniq = new Map();
  for (const r of rows) uniq.set(r.網址, r);
  return [...uniq.values()];
}

function writeXlsx(rows, outPath) {
  // 固定欄位順序：日期、名稱、網址
  const data = rows.map((r) => ({
    日期: r.日期 || "",
    名稱: r.名稱 || "",
    網址: r.網址 || "",
  }));

  const ws = XLSX.utils.json_to_sheet(data, { header: ["日期", "名稱", "網址"] });

  // 欄寬（不影響功能，純粹好看）
  ws["!cols"] = [{ wch: 12 }, { wch: 40 }, { wch: 70 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "posts");
  XLSX.writeFile(wb, outPath);
}

async function main() {
  // ✅ 允許兩種用法：
  // 1) node xxx.mjs               -> 互動輸入網址
  // 2) node xxx.mjs <url>         -> 直接帶網址
  let startUrl = process.argv[2];
  if (!startUrl) startUrl = await askUrl();

  if (!startUrl) {
    console.log("❌ 未輸入網址，結束程式");
    process.exit(1);
  }

  // 自動補 https://
  if (!/^https?:\/\//i.test(startUrl)) {
    startUrl = "https://" + startUrl;
  }

  // 確保輸出資料夾存在
  fs.mkdirSync(OUT_DIR, { recursive: true });

  const headers = { "User-Agent": "Mozilla/5.0" };

  const firstHtml = (await axios.get(startUrl, { headers })).data;
  let $ = cheerio.load(firstHtml);

  const maxPage = getMaxPage($);
  console.log("偵測最大頁碼:", maxPage);

  const seen = new Set();
  const all = [];

  for (let p = 1; p <= maxPage; p++) {
    const url = setPage(startUrl, p);
    const html = p === 1 ? firstHtml : (await axios.get(url, { headers })).data;
    $ = cheerio.load(html);

    const rows = extractPosts($);
    let add = 0;

    for (const r of rows) {
      if (!seen.has(r.網址)) {
        seen.add(r.網址);
        all.push(r);
        add++;
      }
    }

    console.log(`page=${p}/${maxPage} 新增 ${add} 累計 ${all.length}`);
    await sleep(SLEEP_MS);
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
