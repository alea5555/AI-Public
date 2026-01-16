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

/**
 * ✅ 關鍵：清掉 XML 1.0 不允許的控制字元（Word 最常因此要求修復）
 * 允許：\t \n \r，以及 U+0020 以上，排除 surrogate/FFFE/FFFF
 */
function sanitizeXmlText(s) {
  s = (s || "").toString();
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);

    // allow tab, lf, cr
    if (code === 0x09 || code === 0x0A || code === 0x0D) { out += s[i]; continue; }

    // disallow C0 controls
    if (code < 0x20) continue;

    // disallow surrogate halves
    if (code >= 0xD800 && code <= 0xDFFF) continue;

    // disallow non-characters
    if (code === 0xFFFE || code === 0xFFFF) continue;

    out += s[i];
  }
  // 再做一次 whitespace normalize（避免極端空白）
  out = out.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  return out;
}

async function loadSharp() {
  try {
    const mod = await import("sharp");
    return mod.default || mod;
  } catch {
    return null;
  }
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
    .map(sanitizeXmlText)
    .filter(Boolean)
    .filter(s => s.length >= 2)
    .filter(s => !junkContains.some(j => s.includes(j)));

  // 合併碎片
  const merged = [];
  for (const s of filtered) {
    if (!merged.length) { merged.push(s); continue; }
    const prev = merged[merged.length - 1];
    if (prev.length < 12 && s.length < 20 && !/[。！？:：]$/.test(prev)) {
      merged[merged.length - 1] = sanitizeXmlText(`${prev}${s}`);
    } else merged.push(s);
  }

  // 轉段落
  const paras = [];
  let buf = [];
  const flush = () => {
    const t = sanitizeXmlText(buf.join(" "));
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
      } catch { return true; }
    };

    const collectFromRoot = (root) => {
      if (!root) return;

      const elems = root.querySelectorAll ? root.querySelectorAll("h1,h2,h3,h4,p,li,blockquote,pre,span,div") : [];
      for (const el of elems) {
        if (!visible(el)) continue;
        const tag = (el.tagName || "").toLowerCase();
        if (["nav","header","footer","aside","button","input","textarea","select"].includes(tag)) continue;

        const t = norm(el.innerText || el.textContent || "");
        if (t && t.length >= 2 && t.length <= 2000) out.push(t);
      }

      const all = root.querySelectorAll ? root.querySelectorAll("*") : [];
      for (const el of all) {
        if (el.shadowRoot) collectFromRoot(el.shadowRoot);
      }
    };

    collectFromRoot(document);

    const bodyText = norm(document.body?.innerText || "");
    if (bodyText) out.push(bodyText);

    return out;
  });
}

async function collectImages(page) {
  return await page.evaluate(() => {
    const urls = [];
    const push = (u) => { if (u) urls.push(u); };

    const getBgUrls = (el) => {
      try {
        const cs = getComputedStyle(el);
        const bg = cs.backgroundImage || "";
        const matches = [...bg.matchAll(/url\(["']?(.*?)["']?\)/g)];
        return matches.map(m => m[1]).filter(Boolean);
      } catch { return []; }
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
}

async function waitForTextStable(page, { intervalMs = 700, stableTimes = 4, timeoutMs = 45000 } = {}) {
  const start = Date.now();
  let lastLen = -1;
  let stable = 0;

  while (Date.now() - start < timeoutMs) {
    const len = await page.evaluate(() => (document.body?.innerText || "").length);
    if (len === lastLen && len > 200) stable++;
    else stable = 0;
    lastLen = len;
    if (stable >= stableTimes) return len;
    await page.waitForTimeout(intervalMs);
  }
  return lastLen;
}

/**
 * ✅ 關鍵：圖片一律轉 PNG（避免 Word 修復/不支援）
 * - svg 直接跳過（sharp 不一定能轉）
 * - 其他格式：有 sharp 就轉 png；沒 sharp 就只接受 jpeg/png
 */
async function getPngImageBuffer(url, sharp) {
  const { buf, contentType } = await fetchBin(url);
  if (!buf) return null;

  const ct = (contentType || "").toLowerCase();

  // SVG 最容易搞壞 docx/Word，直接跳過
  if (ct.includes("image/svg") || url.toLowerCase().endsWith(".svg")) return null;

  const isPng = ct.includes("image/png") || url.toLowerCase().endsWith(".png");
  const isJpg = ct.includes("image/jpeg") || ct.includes("image/jpg") || /\.jpe?g(\?|#|$)/i.test(url);
  const isWebp = ct.includes("image/webp") || url.toLowerCase().endsWith(".webp");
  const isGif = ct.includes("image/gif") || url.toLowerCase().endsWith(".gif");
  const otherRaster = ct.startsWith("image/") && !isPng && !isJpg;

  // 沒 sharp：只收 jpg/png（其他跳過避免壞檔）
  if (!sharp) {
    if (isPng || isJpg) return buf;
    return null;
  }

  // 有 sharp：統一轉 png（包含 jpg/webp/gif/其他）
  try {
    // gif 取第一幀也可，sharp 基本可處理
    const out = await sharp(buf).png().toBuffer();
    return out;
  } catch {
    // 轉不了就跳過
    return null;
  }
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

    const finalLen = await waitForTextStable(page);
    console.log(`[INFO] body_text_len(stable)=${finalLen}`);

    // 觸發 lazy
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1200);
    await page.evaluate(() => window.scrollTo(0, 0));
    await page.waitForTimeout(600);

    const title = sanitizeXmlText((await page.title()) || "Untitled");

    const shadowLines = await collectShadowText(page);
    const mergedLines = uniqKeepOrder(shadowLines.map(sanitizeXmlText));
    const paras = normalizeTextLines(mergedLines);

    const rawImgs = await collectImages(page);
    const imgUrls = uniqKeepOrder(
      rawImgs.map(u => {
        try { return new URL(u, url).toString(); } catch { return ""; }
      }).filter(Boolean)
    );

    console.log(`[INFO] merged_lines=${mergedLines.length}, paras=${paras.length}, imgs=${imgUrls.length}`);

    if (paras.length < 3) {
      console.log("[ERR] 正文仍不足（可能是 Canvas 文字）。需要 OCR 才能保證。");
      return;
    }

    const sharp = await loadSharp();

    // ✅ 不放任何超連結，避免 Word 連結/修復提示
    const children = [];
    children.push(new Paragraph({ text: title, heading: HeadingLevel.TITLE }));
    children.push(new Paragraph({ children: [new TextRun({ text: sanitizeXmlText(`來源：${url}`) })] }));
    children.push(new Paragraph({ text: "" }));

    for (const p0 of paras) {
      const p = sanitizeXmlText(p0);
      if (!p) continue;

      const isHeading =
        (p.length <= 28 && p.includes("：")) ||
        /^([一二三四五六七八九十]、|[0-9]+[.)])/.test(p);

      if (isHeading) children.push(new Paragraph({ text: p, heading: HeadingLevel.HEADING_2 }));
      else children.push(new Paragraph({ text: p }));
    }

    // 圖片：全部轉成 PNG 再內嵌
    if (imgUrls.length) {
      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "圖片", heading: HeadingLevel.HEADING_2 }));

      for (const u of imgUrls) {
        const pngBuf = await getPngImageBuffer(u, sharp);
        if (!pngBuf) continue;

        children.push(new Paragraph({
          children: [new ImageRun({ data: pngBuf, transformation: { width: 600, height: 380 } })],
        }));

        await new Promise(r => setTimeout(r, 120));
      }
    }

    const doc = new Document({ sections: [{ children }] });
    const outBuf = await Packer.toBuffer(doc);

    const outName = `${safeFilename(title)}_${tsName()}.docx`;
    const outPath = path.join(process.cwd(), outName);
    fs.writeFileSync(outPath, outBuf);

    console.log(`[OK] 已輸出：${outPath}`);
    if (!sharp) console.log("[TIP] 沒裝 sharp：只會內嵌 jpg/png，其他格式會跳過（但不會壞檔）。");
  } catch (e) {
    console.error("[ERR]", e?.message || e);
  } finally {
    await context.close();
  }
}

main();
