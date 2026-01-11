import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";

const SLEEP_MS = 800;

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
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
    try {
      const u = new URL(href, "https://dummy.base/");
      const p = u.searchParams.get("page");
      if (p && /^\d+$/.test(p)) maxPage = Math.max(maxPage, parseInt(p, 10));
    } catch {}
    const txt = $(el).text().trim();
    if (/^\d+$/.test(txt)) maxPage = Math.max(maxPage, parseInt(txt, 10));
  });

  return maxPage;
}

function extractPosts($) {
  const rows = [];
  let articles = $("article");
  if (articles.length === 0) articles = $("h2");

  articles.each((_, el) => {
    const a = $(el).is("h2") ? $(el).find("a").first() : $(el).find("h2 a").first();
    const link = a.attr("href");
    const title = a.text().trim();
    if (!link || !title) return;

    let dt = "";
    const timeEl = $(el).find("time[datetime]").first();
    if (timeEl.length) dt = (timeEl.attr("datetime") || "").slice(0, 10);

    rows.push({ 日期: dt, 名稱: title, 網址: link.trim() });
  });

  // 去重
  const uniq = new Map();
  for (const r of rows) uniq.set(r.網址, r);
  return [...uniq.values()];
}

async function main() {
  const startUrl = process.argv[2];
  if (!startUrl) {
    console.log('用法：node pixnet_all.js "https://ey90223.pixnet.net/blog"');
    process.exit(1);
  }

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

  fs.writeFileSync("pixnet_posts.json", JSON.stringify(all, null, 2), "utf-8");
  console.log("輸出：pixnet_posts.json");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
