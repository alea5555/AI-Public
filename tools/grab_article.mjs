// grab_article.mjs
import fs from "fs";
import path from "path";
import readline from "readline";

let chromium;
try {
  ({ chromium } = await import("playwright"));
} catch (e) {
  console.error("找不到 playwright。若你尚未安裝，請在專案目錄執行：npm i playwright");
  process.exit(1);
}

function ask(q) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(q, (ans) => { rl.close(); resolve(ans.trim()); }));
}

function tsName() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

async function extractMainText(page) {
  // 等「主要內容」比較可能出現
  // networkidle 有時對 SPA 不靈，所以混合用：domcontentloaded + 少量 wait
  await page.waitForLoadState("domcontentloaded");
  await page.waitForTimeout(1200);

  // 有些站會滾一下才補齊 lazy content
  try {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(800);
    await page.evaluate(() => window.scrollTo(0, 0));
    await page.waitForTimeout(400);
  } catch {}

  const result = await page.evaluate(() => {
    const norm = (s) => (s || "")
      .replace(/\u00a0/g, " ")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();

    const pickCandidates = () => {
      const sels = [
        "article",
        "main",
        "[class*='prose']",
        "[class*='content']",
        "[class*='article']",
        "#content",
        "body"
      ];
      const set = new Set();
      const arr = [];
      for (const sel of sels) {
        document.querySelectorAll(sel).forEach(el => {
          if (!set.has(el)) { set.add(el); arr.push(el); }
        });
      }
      return arr;
    };

    const isJunk = (el) => {
      const tag = (el.tagName || "").toLowerCase();
      if (["nav","header","footer","aside","script","style","noscript"].includes(tag)) return true;
      const cls = (el.className || "").toString().toLowerCase();
      const id = (el.id || "").toLowerCase();
      const bad = ["sidebar","comment","breadcrumb","related","recommend","share","ads","advert","footer"];
      return bad.some(k => cls.includes(k) || id.includes(k));
    };

    const textLen = (el) => {
      if (!el || isJunk(el)) return 0;
      const t = (el.innerText || "").trim();
      return t.length;
    };

    const score = (el) => {
      if (!el || isJunk(el)) return -1e9;
      const t = (el.innerText || "").trim();
      const len = t.length;

      const p = el.querySelectorAll("p").length;
      const li = el.querySelectorAll("li").length;
      const h = el.querySelectorAll("h1,h2,h3,h4").length;
      const pre = el.querySelectorAll("pre").length;
      const bq = el.querySelectorAll("blockquote").length;
      const img = el.querySelectorAll("img").length;

      // 基本：越長越好，但不要無限加
      let s = Math.min(len, 20000);

      // 結構加權：有段落/清單/標題更像文章
      s += p * 250 + li * 120 + h * 180 + pre * 120 + bq * 150 + img * 10;

      // 懲罰 junk 子節點
      const junk = el.querySelectorAll("nav,header,footer,aside").length;
      s -= junk * 800;

      return s;
    };

    const best = (() => {
      const cands = pickCandidates();
      let bestEl = null;
      let bestScore = -1e18;
      for (const el of cands) {
        const sc = score(el);
        if (sc > bestScore) { bestScore = sc; bestEl = el; }
      }
      return bestEl || document.body;
    })();

    const title = norm(document.title || "");
    const mainText = norm(best.innerText || "");

    // 若抽到的 mainText 太短，再退回 body
    if (mainText.length < 200) {
      return { title, mainText: norm(document.body?.innerText || ""), used: "body_fallback" };
    }
    return { title, mainText, used: "best_container" };
  });

  return result;
}

async function main() {
  const url = await ask("請輸入文章網址：\n");
  if (!url) {
    console.log("未輸入網址，結束。");
    return;
  }

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36",
    locale: "zh-TW",
  });
  const page = await context.newPage();

  try {
    console.log("[INFO] Loading...");
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });

    const { title, mainText, used } = await extractMainText(page);

    const out = [
      `Title: ${title}`,
      `URL: ${url}`,
      `Extractor: ${used}`,
      "",
      mainText
    ].join("\n");

    console.log("\n===== 抓取結果（前 2000 字） =====\n");
    console.log(out.slice(0, 2000));
    console.log("\n===== 結束（已完整存檔） =====\n");

    const outName = `output_${tsName()}.txt`;
    const outPath = path.join(process.cwd(), outName);
    fs.writeFileSync(outPath, out, "utf-8");
    console.log(`[OK] Saved: ${outPath}`);
  } catch (e) {
    console.error("[ERR]", e?.message || e);
  } finally {
    await context.close();
    await browser.close();
  }
}

main();
