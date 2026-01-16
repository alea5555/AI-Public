import fs from "fs";
import path from "path";
import readline from "readline";
import { chromium } from "playwright";
import { Document, Packer, Paragraph, HeadingLevel, TextRun, ImageRun } from "docx";
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

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
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function sanitizeXmlText(s) {
  s = (s || "").toString();
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code === 0x09 || code === 0x0A || code === 0x0D) { out += s[i]; continue; }
    if (code < 0x20) continue;
    if (code >= 0xD800 && code <= 0xDFFF) continue;
    if (code === 0xFFFE || code === 0xFFFF) continue;
    out += s[i];
  }
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

  const merged = [];
  for (const s of filtered) {
    if (!merged.length) { merged.push(s); continue; }
    const prev = merged[merged.length - 1];
    if (prev.length < 12 && s.length < 20 && !/[。！？:：]$/.test(prev)) {
      merged[merged.length - 1] = sanitizeXmlText(`${prev}${s}`);
    } else merged.push(s);
  }

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
    const norm = (s) => (s || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();

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
        if (t && t.length >= 2 && t.length <= 4000) out.push(t);
      }
      const all = root.querySelectorAll ? root.querySelectorAll("*") : [];
      for (const el of all) if (el.shadowRoot) collectFromRoot(el.shadowRoot);
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

async function getPngImageBuffer(url, sharp) {
  const { buf, contentType } = await fetchBin(url);
  if (!buf) return null;

  const ct = (contentType || "").toLowerCase();
  const lowerUrl = url.toLowerCase();

  // svg 直接跳過（最容易搞出 Word 問題）
  if (ct.includes("image/svg") || lowerUrl.endsWith(".svg")) return null;

  const isPng = ct.includes("image/png") || lowerUrl.endsWith(".png");
  const isJpg = ct.includes("image/jpeg") || ct.includes("image/jpg") || /\.jpe?g(\?|#|$)/i.test(url);

  if (!sharp) {
    // 沒 sharp 就只收 jpg/png（避免壞檔）
    if (isPng || isJpg) return buf;
    return null;
  }

  try {
    return await sharp(buf).png().toBuffer();
  } catch {
    return null;
  }
}

/**
 * ✅ 封包層清洗：刪掉所有 External relationships（Word 看到這個就會跳）
 */
async function stripExternalRelationships(docxBuffer) {
  const zip = await JSZip.loadAsync(docxBuffer);
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", suppressEmptyNode: true });

  // 只要在 word/_rels/*.rels 都掃
  const relPaths = Object.keys(zip.files).filter(p => /^word\/_rels\/.*\.rels$/i.test(p));
  for (const p of relPaths) {
    const xml = await zip.file(p).async("string");
    let obj;
    try { obj = parser.parse(xml); } catch { continue; }

    const rels = obj?.Relationships?.Relationship;
    if (!rels) continue;

    const arr = Array.isArray(rels) ? rels : [rels];

    const kept = arr.filter(r => {
      const tm = (r["@_TargetMode"] || "").toString();
      const target = (r["@_Target"] || "").toString();
      const t = target.toLowerCase();
      // 刪掉 external 或 任何看起來是網路/檔案路徑的 target
      if (tm.toLowerCase() === "external") return false;
      if (t.startsWith("http://") || t.startsWith("https://") || t.startsWith("file:")) return false;
      return true;
    });

    obj.Relationships.Relationship = kept.length === 1 ? kept[0] : kept;
    const newXml = builder.build(obj);
    zip.file(p, newXml);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
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

    console.log(`[INFO] paras=${paras.length}, imgs=${imgUrls.length}`);

    if (paras.length < 3) {
      console.log("[ERR] 正文仍不足（可能是 Canvas 文字）。需要 OCR 才能保證。");
      return;
    }

    const sharp = await loadSharp();

    // ✅ 不建立超連結：來源只放純文字（而且把 https:// 變成 hxxps:// 避免任何自動判定）
    const safeUrlText = sanitizeXmlText(url.replace(/^https?:\/\//i, "hxxps://"));

    const children = [];
    children.push(new Paragraph({ text: title, heading: HeadingLevel.TITLE }));
    children.push(new Paragraph({ children: [new TextRun({ text: `來源：${safeUrlText}` })] }));
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

    if (imgUrls.length) {
      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "圖片", heading: HeadingLevel.HEADING_2 }));

      for (const u of imgUrls) {
        const pngBuf = await getPngImageBuffer(u, sharp);
        if (!pngBuf) continue;

        children.push(new Paragraph({
          children: [new ImageRun({ data: pngBuf, transformation: { width: 600, height: 380 } })],
        }));
      }
    }

    const doc = new Document({ sections: [{ children }] });
    let outBuf = await Packer.toBuffer(doc);

    // ✅ 最終保險：刪掉封包中所有 External relationships（Word 就不會跳）
    outBuf = await stripExternalRelationships(outBuf);

    const outName = `${safeFilename(title)}_${tsName()}.docx`;
    const outPath = path.join(process.cwd(), outName);
    fs.writeFileSync(outPath, outBuf);

    console.log(`[OK] 已輸出：${outPath}`);
  } catch (e) {
    console.error("[ERR]", e?.message || e);
  } finally {
    await context.close();
  }
}

main();
