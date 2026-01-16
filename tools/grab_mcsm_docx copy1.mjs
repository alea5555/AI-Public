import fs from "fs";
import path from "path";
import readline from "readline";
import { chromium } from "playwright";
import { Document, Packer, Paragraph, HeadingLevel, TextRun, ImageRun } from "docx";

function ask(q) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(q, (ans) => { rl.close(); resolve((ans || "").trim()); }));
}

function safeFilename(s, maxLen = 120) {
  s = (s || "").trim();
  if (!s) return "output";
  s = s.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
  s = s.replace(/\s+/g, " ").trim();
  if (s.length > maxLen) s = s.slice(0, maxLen).trim();
  return s || "output";
}

function tsName() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function absUrl(base, u) {
  try { return new URL(u, base).toString(); } catch { return ""; }
}

async function loadSharp() {
  try {
    const mod = await import("sharp");
    return mod.default || mod;
  } catch {
    return null;
  }
}

function isWebp(url, contentType) {
  const u = (url || "").toLowerCase();
  if ((contentType || "").includes("image/webp")) return true;
  return u.endsWith(".webp");
}

async function fetchBin(url) {
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) return { buf: null, contentType: "" };
  const contentType = (res.headers.get("content-type") || "").toLowerCase();
  const ab = await res.arrayBuffer();
  return { buf: Buffer.from(ab), contentType };
}

function uniqKeepOrder(arr) {
  const seen = new Set();
  const out = [];
  for (const x of arr) {
    const k = (x || "").trim();
    if (!k) continue;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(k);
  }
  return out;
}

function normalizeTextLines(lines) {
  const junkContains = ["MULTISTRATEGYS", "策略市集", "策略生成器"];
  const filtered = lines
    .map(s => s.replace(/\u00a0/g, " ").trim())
    .filter(Boolean)
    .filter(s => s.length >= 2)
    .filter(s => !junkContains.some(j => s.includes(j)));

  // 合併碎片
  const merged = [];
  for (const s of filtered) {
    if (!merged.length) { merged.push(s); continue; }
    const prev = merged[merged.length - 1];
    if (prev.length < 12 && s.length < 20 && !/[。！？:：]$/.test(prev)) {
      merged[merged.length - 1] = `${prev}${s}`;
    } else merged.push(s);
  }

  // 轉段落
  const paras = [];
  let buf = [];
  const flush = () => {
    const t = buf.join(" ").replace(/\s+/g, " ").trim();
    if (t) paras.push(t);
    buf = [];
  };

  for (const s of merged) {
    const isHeadingLike =
      /^([一二三四五六七八九十]、|[0-9]+[.)])/.test(s) ||
      (s.length <= 28 && /：/.test(s));

    if (isHeadingLike && buf.length) flush();

    buf.push(s);

    const cur = buf.join(" ");
    if (/[。！？]$/.test(s) || cur.length > 260) flush();
  }
  flush();

  return paras;
}

async function collectShadowText(page) {
  // ✅ 只靠 DOM + ShadowRoot 深度走訪，抓可見文字
  return await page.evaluate(() => {
    const out = [];
    const norm = (s) => (s || "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    const visible = (el) => {
      try {
        const cs = getComputedStyle(el);
        if (cs.display === "none" || cs.visibility === "hidden" || cs.opacity === "0") return false;
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) return false;
        return true;
      } catch {
        return true;
      }
    };

    const collectFromRoot = (root) => {
      if (!root) return;

      // 先抓語意元素（最像文章）
      const elems = root.querySelectorAll ? root.querySelectorAll("h1,h2,h3,h4,p,li,blockquote,pre,span,div") : [];
      for (const el of elems) {
        if (!visible(el)) continue;

        // 避免導航/按鈕
        const tag = (el.tagName || "").toLowerCase();
        if (["nav","header","footer","aside","button","input","textarea","select"].includes(tag)) continue;

        const t = norm(el.innerText || el.textContent || "");
        if (t && t.length >= 2 && t.length <= 2000) out.push(t);
      }

      // 走訪 shadow
      const all = root.querySelectorAll ? root.querySelectorAll("*") : [];
      for (const el of all) {
        if (el.shadowRoot) collectFromRoot(el.shadowRoot);
      }
    };

    collectFromRoot(document);

    // 再補：整頁可見文字（保底）
    const bodyText = norm(document.body?.innerText || "");
    if (bodyText) out.push(bodyText);

    return out;
  });
}

async function collectImages(page, baseUrl) {
  const imgs = await page.evaluate(() => {
    const urls = [];
    const push = (u) => { if (u) urls.push(u); };

    const getBgUrls = (el) => {
      try {
        const cs = getComputedStyle(el);
        const bg = cs.backgroundImage || "";
        const matches = [...bg.matchAll(/url\(["']?(.*?)["']?\)/g)];
        return matches.map(m => m[1]).filter(Boolean);
      } catch {
        return [];
      }
    };

    const walk = (root) => {
      if (!root) return;
      if (root.shadowRoot) walk(root.shadowRoot);

      const nodes = root.querySelectorAll ? root.querySelectorAll("*") : [];
      for (const el of nodes) {
        const tag = (el.tagName || "").toLowerCase();
        if (tag === "img") {
          push(el.currentSrc || el.getAttribute("src") || el.getAttribute("data-src") || el.getAttribute("data-original") || el.getAttribute("data-lazy-src"));
        }
        for (const u of getBgUrls(el)) push(u);
        if (el.shadowRoot) walk(el);
      }
    };

    walk(document);
    return urls;
  });

  const abs = imgs.map(u => {
    try { return new URL(u, baseUrl).toString(); } catch { return ""; }
  }).filter(Boolean);

  return uniqKeepOrder(abs);
}

async function main() {
  const url = await ask("請輸入文章網址：\n");
  if (!url) return console.log("未輸入網址，結束。");

  const userDataDir = path.join(process.cwd(), "pw_profile_mcsm");
  const context = await chromium.launchPersistentContext(userDataDir, {
    headless: false,
    viewport: { width: 1400, height: 900 },
    locale: "zh-TW",
    args: ["--disable-blink-features=AutomationControlled"],
  });

  const page = await context.newPage();

  try {
    console.log("[INFO] 打開瀏覽器中...");
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

    console.log("✅ 請在瀏覽器視窗確認『正文已完整顯示』，並可捲動到內文。");
    console.log("👉 確認好後回到 CMD 按 Enter 開始輸出 docx...");
    await ask("");

    // 觸發 lazy
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1200);
    await page.evaluate(() => window.scrollTo(0, 0));
    await page.waitForTimeout(600);

    const title = (await page.title()) || "Untitled";

    // ✅ 抓正文（shadow 深度走訪）
    const shadowLines = await collectShadowText(page);
    const mergedLines = uniqKeepOrder(shadowLines);

    const paras = normalizeTextLines(mergedLines);

    // 圖片
    const imgUrls = await collectImages(page, url);

    console.log(`[INFO] merged_lines=${mergedLines.length}, paras=${paras.length}, imgs=${imgUrls.length}`);

    if (paras.length < 3) {
      console.log("[ERR] 仍抓不到足夠正文：這頁很可能用 Canvas/圖片方式渲染文字。");
      console.log("      下一步只能 OCR 截圖硬解（我可以直接給你 OCR 版）。");
      return;
    }

    const sharp = await loadSharp();

    // 產 docx
    const children = [];
    children.push(new Paragraph({ text: title, heading: HeadingLevel.TITLE }));
    children.push(new Paragraph({ children: [new TextRun({ text: `來源：${url}` })] }));
    children.push(new Paragraph({ text: "" }));

    for (const p of paras) {
      const isHeading =
        (p.length <= 28 && p.includes("：")) ||
        /^([一二三四五六七八九十]、|[0-9]+[.)])/.test(p);
      if (isHeading) children.push(new Paragraph({ text: p, heading: HeadingLevel.HEADING_2 }));
      else children.push(new Paragraph({ text: p }));
    }

    if (imgUrls.length) {
      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "圖片", heading: HeadingLevel.HEADING_2 }));

      for (const u of imgUrls) {
        const { buf, contentType } = await fetchBin(u);
        if (!buf) continue;

        let imgBuf = buf;
        if (isWebp(u, contentType)) {
          if (!sharp) continue;
          try { imgBuf = await sharp(buf).png().toBuffer(); } catch { continue; }
        }

        children.push(new Paragraph({
          children: [new ImageRun({ data: imgBuf, transformation: { width: 600, height: 380 } })],
        }));

        await new Promise(r => setTimeout(r, 200));
      }
    }

    const doc = new Document({ sections: [{ children }] });
    const outBuf = await Packer.toBuffer(doc);

    const outName = `${safeFilename(title)}_${tsName()}.docx`;
    const outPath = path.join(process.cwd(), outName);
    fs.writeFileSync(outPath, outBuf);

    console.log(`[OK] 已輸出：${outPath}`);
    if (!sharp) console.log("[TIP] 你沒裝 sharp，webp 圖片會被略過；建議 npm i sharp 再跑一次。");
  } catch (e) {
    console.error("[ERR]", e?.message || e);
  } finally {
    await context.close();
  }
}

main();
