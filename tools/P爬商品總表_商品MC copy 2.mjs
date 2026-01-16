// 檔名：P爬商品總表_商品MC.mjs
// 用法：node P爬商品總表_商品MC.mjs
//       然後輸入 https://mcsm.multicharts.com.tw/ 或 https://mcsm.multicharts.com.tw/product/info/1

import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import readline from "readline";
import XLSX from "xlsx";

const HEADERS = { "User-Agent": "Mozilla/5.0" };

// 速度與穩定性參數
const SLEEP_MS = 450;           // 每次請求間隔（不要太快）
const SAVE_EVERY = 10;          // 每新增幾筆就存檔一次（避免中斷白做）
const HARD_MAX_ID = 20000;      // 絕對保險上限（防止無限掃）
const GAP_LIMIT = 500;          // 距離「最後一次成功」超過這個範圍還沒成功，就判斷真的沒有了

// 上限探測（加速用）
const PROBE_BLOCK = 200;        // 探測時一次檢查的區塊大小
const PROBE_STEP_START = 200;   // 指數跳躍的起始步長
const PROBE_MAX_ROUNDS = 20;    // 最多跳幾次（保險）

const OUT_XLSX = "product_total.xlsx";
const OUT_CSV = "product_total.csv";

/* ------------------ 小工具 ------------------ */

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function askUrl() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question("請輸入第一個商品網址或根網址（例如 https://mcsm.multicharts.com.tw/ 或 .../product/info/1）：\n", (ans) => {
      rl.close();
      resolve((ans || "").trim());
    });
  });
}

function toAbs(url) {
  if (!url) return "";
  if (/^https?:\/\//i.test(url)) return url;
  return "https://" + url;
}

function normalizeId(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function buildTemplateFromProductUrl(productUrl) {
  // 允許：https://xxx/product/info/25
  const u = new URL(productUrl);
  const parts = u.pathname.split("/").filter(Boolean);
  const last = parts.at(-1);
  if (!/^\d+$/.test(last)) throw new Error("網址最後一段必須是數字 ID");
  const startId = parseInt(last, 10);
  const basePath = parts.slice(0, -1).join("/");
  const template = `${u.origin}/${basePath}/{id}`;
  return { startId, template, origin: u.origin };
}

function buildTemplateFromRoot(origin) {
  // 固定規則：/product/info/{id}
  const u = new URL(origin);
  return { template: `${u.origin}/product/info/{id}`, origin: u.origin };
}

function firstMatchNumber(str) {
  const m = String(str || "").match(/-?\d[\d,]*(?:\.\d+)?/);
  return m ? m[0].replace(/,/g, "") : "";
}

/* ------------------ 解析「根網址→第一個商品」 ------------------ */

async function findFirstProductFromRoot(rootUrl) {
  const res = await axios.get(rootUrl, { headers: HEADERS, timeout: 20000 });
  const $ = cheerio.load(res.data);

  // 找到第一個 /product/info/{id}
  let best = null;

  $("a[href]").each((_, a) => {
    const href = $(a).attr("href");
    if (!href) return;

    const m = href.match(/\/product\/info\/(\d+)/);
    if (!m) return;

    const id = parseInt(m[1], 10);
    if (!Number.isFinite(id)) return;

    if (best == null || id < best) best = id;
  });

  if (best == null) {
    throw new Error("在首頁找不到任何 /product/info/{id} 連結");
  }

  const { template } = buildTemplateFromRoot(rootUrl);
  return { startId: best, template };
}

/* ------------------ 抓單一商品（含欄位） ------------------ */

async function fetchProduct(id, tpl) {
  const url = tpl.replace("{id}", String(id));

  try {
    const res = await axios.get(url, { headers: HEADERS, timeout: 20000 });
    const html = res.data;
    const $ = cheerio.load(html);

    // 判斷是否真有商品：通常會有 h1，或 og:title
    const title =
      $("h1").first().text().trim() ||
      $('meta[property="og:title"]').attr("content")?.trim() ||
      "";

    if (!title) return null;

    // 老師：盡量抓卡片中顯示的名字（不保證每頁一樣）
    let teacher = "";
    // 常見：有個人 icon 後面文字、或某些 class；抓不到就留空
    const teacherCandidate = $("a, span, div")
      .filter((_, el) => {
        const t = $(el).text().trim();
        return t.endsWith("老師") && t.length <= 20;
      })
      .first()
      .text()
      .trim();
    if (teacherCandidate) teacher = teacherCandidate;

    // 商品代碼/標籤（例如 TXF1、台指期 等）
    const tags = [];
    // 抓 chip/標籤類文字：常見是一些小方塊
    $("a, span, div").each((_, el) => {
      const t = $(el).text().trim();
      if (!t) return;
      // 排除太長的段落
      if (t.length > 18) return;
      // 過濾一些常見非標籤字
      if (t === "查看詳情" || t === "策略市集") return;
      // 常見商品/標籤格式
      if (/^(#)?[A-Z]{2,6}\d?$/.test(t) || /台指|加權|櫃買|期貨|策略|程式交易|自動交易/.test(t)) {
        tags.push(t.replace(/^#/, ""));
      }
    });

    const uniqTags = [...new Set(tags)].slice(0, 12);
    const tagText = uniqTags.join(" / ");

    // 以全文關鍵字找數值（抓不到就空白）
    const pageText = $("body").text().replace(/\s+/g, " ");

    const profit = (() => {
      // 淨利 +x,xxx,xxx
      const idx = pageText.indexOf("淨利");
      if (idx < 0) return "";
      return firstMatchNumber(pageText.slice(idx, idx + 80));
    })();

    const winRate = (() => {
      const idx = pageText.indexOf("勝率");
      if (idx < 0) return "";
      return firstMatchNumber(pageText.slice(idx, idx + 80));
    })();

    const rr = (() => {
      const idx = pageText.indexOf("風報比");
      if (idx < 0) return "";
      return firstMatchNumber(pageText.slice(idx, idx + 80));
    })();

    return {
      ID: id,
      名稱: title,
      老師: teacher,
      淨利: profit,
      勝率: winRate,
      風報比: rr,
      標籤: tagText,
      網址: url,
    };
  } catch {
    return null;
  }
}

/* ------------------ 斷點續抓：讀舊檔 ------------------ */

function loadExistingRows() {
  if (!fs.existsSync(OUT_XLSX)) return [];

  try {
    const wb = XLSX.readFile(OUT_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws);

    // 統一欄位名並確保 ID 是數字
    const cleaned = [];
    for (const r of rows) {
      const id = normalizeId(r.ID ?? r.Id ?? r.id);
      if (id == null) continue;

      cleaned.push({
        ID: id,
        名稱: String(r.名稱 ?? r.Title ?? r.title ?? "").trim(),
        老師: String(r.老師 ?? r.Teacher ?? r.teacher ?? "").trim(),
        淨利: String(r.淨利 ?? r.Profit ?? r.profit ?? "").trim(),
        勝率: String(r.勝率 ?? r.WinRate ?? r.winrate ?? "").trim(),
        風報比: String(r.風報比 ?? r.RR ?? r.rr ?? "").trim(),
        標籤: String(r.標籤 ?? r.Tags ?? r.tags ?? "").trim(),
        網址: String(r.網址 ?? r.URL ?? r.url ?? "").trim(),
      });
    }

    return cleaned;
  } catch {
    return [];
  }
}

/* ------------------ 輸出：XLSX + CSV ------------------ */

function writeOutputs(rows) {
  const headers = ["ID", "名稱", "老師", "淨利", "勝率", "風報比", "標籤", "網址"];

  // XLSX
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  ws["!cols"] = [
    { wch: 8 },   // ID
    { wch: 50 },  // 名稱
    { wch: 18 },  // 老師
    { wch: 14 },  // 淨利
    { wch: 10 },  // 勝率
    { wch: 10 },  // 風報比
    { wch: 35 },  // 標籤
    { wch: 70 },  // 網址
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "products");
  XLSX.writeFile(wb, OUT_XLSX);

  // CSV
  const escapeCsv = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv =
    headers.join(",") +
    "\n" +
    rows.map((r) => headers.map((h) => escapeCsv(r[h])).join(",")).join("\n");
  fs.writeFileSync(OUT_CSV, csv, "utf8");
}

/* ------------------ ① 探測合理上界（加速） ------------------ */

async function blockHasAnySuccess(fromId, toId, tpl, doneSet) {
  // 探測用：在 [fromId, toId] 之間找「是否存在任一有效商品」
  // 為了加速：每 10 個取樣一次 + 端點，避免整段全掃
  if (toId < fromId) return false;

  const candidates = new Set();
  candidates.add(fromId);
  candidates.add(toId);

  const step = 10;
  for (let i = fromId; i <= toId; i += step) candidates.add(i);

  const sorted = [...candidates].sort((a, b) => a - b);

  for (const id of sorted) {
    if (doneSet?.has(id)) return true; // 已抓過表示一定有效
    const r = await fetchProduct(id, tpl);
    await sleep(Math.min(200, SLEEP_MS));
    if (r) return true;
  }
  return false;
}

async function probeUpperBound(startId, tpl, doneSet) {
  // 指數跳躍：找一個「之後一大段都沒有資料」的合理上界
  let lastSeen = startId;
  let step = PROBE_STEP_START;

  for (let round = 0; round < PROBE_MAX_ROUNDS; round++) {
    const hi = Math.min(lastSeen + step, HARD_MAX_ID);
    const from = Math.max(1, hi - PROBE_BLOCK + 1);

    const ok = await blockHasAnySuccess(from, hi, tpl, doneSet);
    if (ok) {
      lastSeen = hi;
      step *= 2;
      if (hi >= HARD_MAX_ID) break;
    } else {
      // 找到一個「區塊內看起來都沒有」的位置，當成上界候選
      return hi;
    }
  }

  return Math.min(lastSeen + PROBE_BLOCK, HARD_MAX_ID);
}

/* ------------------ 主流程：②③④ 全整合 ------------------ */

async function main() {
  let input = process.argv[2];
  if (!input) input = await askUrl();
  input = toAbs(input);

  let startId, template;

  // 允許三種輸入：
  // 1) 根網址 https://xxx/
  // 2) 商品頁 https://xxx/product/info/1
  // 3) 任何包含 /product/info/ 的網址（最後段是數字）
  try {
    const parsed = buildTemplateFromProductUrl(input);
    startId = parsed.startId;
    template = parsed.template;
  } catch {
    console.log("ℹ 偵測為根網址/非數字結尾，嘗試自動尋找第一個商品…");
    const found = await findFirstProductFromRoot(input);
    startId = found.startId;
    template = found.template;
  }

  console.log(`✔ 連號模板：${template}`);
  console.log(`✔ 起始 ID：${startId}`);

  // ④ 斷點續抓
  const existing = loadExistingRows();
  const rows = [...existing];

  const done = new Set(rows.map((r) => normalizeId(r.ID)).filter((x) => x != null));
  console.log(`✔ 已有舊檔：${rows.length} 筆（將自動續抓、跳過已完成 ID）`);

  // ① 探測合理上界（加速）
  console.log("🔍 探測合理上界中（加速用）…");
  const upperGuess = await probeUpperBound(startId, template, done);
  console.log(`✔ 探測上界候選：ID ≈ ${upperGuess}`);
  console.log("🚀 開始正式掃描（不中斷、不怕斷號）…\n");

  let lastSuccessId = 0;
  let addedSinceSave = 0;

  // 正式掃描：從 startId 起掃到 upperGuess，再用 GAP_LIMIT 規則延伸，避免漏後段
  // 做法：先掃到 upperGuess，然後再延伸掃到 (lastSuccessId + GAP_LIMIT) 為止。
  let id = startId;
  let hardStop = false;

  while (!hardStop) {
    if (id > HARD_MAX_ID) break;

    // 若已抓過就跳過，但要更新 lastSuccessId（因為這代表成功）
    if (done.has(id)) {
      lastSuccessId = Math.max(lastSuccessId, id);
      id++;
      continue;
    }

    const r = await fetchProduct(id, template);
    if (r) {
      rows.push(r);
      done.add(id);
      lastSuccessId = id;
      addedSinceSave++;

      // 顯示
      console.log(`✔ ID=${id} ${r.名稱}`);
    }

    // ③ 不漏後段的停止邏輯（核心）
    // 一開始要先跑到 upperGuess（探測上界只是加速，不是結尾）
    // 之後：只要距離最後成功超過 GAP_LIMIT 都沒成功，就停
    if (id >= upperGuess && lastSuccessId > 0 && id - lastSuccessId >= GAP_LIMIT) {
      console.log(`\n🛑 超過最後成功 ID=${lastSuccessId} 後 ${GAP_LIMIT} 筆皆無資料，停止`);
      break;
    }

    // 定期存檔（④ 防中斷）
    if (addedSinceSave >= SAVE_EVERY) {
      // 依 ID 排序、去重保險
      const uniq = new Map();
      for (const x of rows) uniq.set(Number(x.ID), x);
      const sorted = [...uniq.values()].sort((a, b) => Number(a.ID) - Number(b.ID));
      rows.length = 0;
      rows.push(...sorted);

      writeOutputs(rows);
      console.log(`💾 已自動存檔（目前 ${rows.length} 筆）`);
      addedSinceSave = 0;
    }

    id++;
    await sleep(SLEEP_MS);
  }

  // 最終輸出
  const uniq = new Map();
  for (const x of rows) uniq.set(Number(x.ID), x);
  const sorted = [...uniq.values()].sort((a, b) => Number(a.ID) - Number(b.ID));

  writeOutputs(sorted);

  console.log(`\n✅ 完成`);
  console.log(`📦 共 ${sorted.length} 筆`);
  console.log(`📄 輸出：${OUT_XLSX}`);
  console.log(`📄 輸出：${OUT_CSV}`);
}

main().catch((e) => {
  console.error("❌", e?.message || e);
  process.exit(1);
});
