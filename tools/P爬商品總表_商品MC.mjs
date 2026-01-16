// 檔名：P爬商品總表_商品MC.mjs
import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import readline from "readline";
import XLSX from "xlsx";

const HEADERS = { "User-Agent": "Mozilla/5.0" };

// ====== 速度/穩定性參數 ======
const SLEEP_MS = 450;            // 每次請求間隔
const SAVE_EVERY = 10;           // 每新增幾筆就自動存檔
const HARD_MAX_ID = 20000;       // 絕對保險上限（防止無限掃）
const GAP_LIMIT = 2000;          // 距離最後成功超過這個範圍仍無成功 → 結束

// ====== 心跳（你要的） ======
const HEARTBEAT_SEC = 10;        // 每 10 秒印一次進度（一定會印）
const SHOW_MISS_EVERY = 100;     // 連續空號每 100 個印一次（避免刷屏）

// ====== 上限探測（加速用） ======
const PROBE_BLOCK = 200;
const PROBE_STEP_START = 200;
const PROBE_MAX_ROUNDS = 20;

const OUT_XLSX = "product_total.xlsx";
const OUT_CSV = "product_total.csv";

/* ------------------ utils ------------------ */

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

function nowStamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function firstMatchNumber(str) {
  const m = String(str || "").match(/-?\d[\d,]*(?:\.\d+)?/);
  return m ? m[0].replace(/,/g, "") : "";
}

function buildTemplateFromProductUrl(productUrl) {
  const u = new URL(productUrl);
  const parts = u.pathname.split("/").filter(Boolean);
  const last = parts.at(-1);
  if (!/^\d+$/.test(last)) throw new Error("網址最後一段必須是數字 ID");
  const startId = parseInt(last, 10);
  const basePath = parts.slice(0, -1).join("/");
  const template = `${u.origin}/${basePath}/{id}`;
  return { startId, template };
}

async function findFirstProductFromRoot(rootUrl) {
  const res = await axios.get(rootUrl, { headers: HEADERS, timeout: 20000 });
  const $ = cheerio.load(res.data);

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

  if (best == null) throw new Error("在首頁找不到任何 /product/info/{id} 連結");
  const u = new URL(rootUrl);
  return { startId: best, template: `${u.origin}/product/info/{id}` };
}

/* ------------------ fetch product ------------------ */

async function fetchProduct(id, tpl) {
  const url = tpl.replace("{id}", String(id));
  try {
    const res = await axios.get(url, { headers: HEADERS, timeout: 20000 });
    const $ = cheerio.load(res.data);

    const title =
      $("h1").first().text().trim() ||
      $('meta[property="og:title"]').attr("content")?.trim() ||
      "";
    if (!title) return null;

    let teacher = "";
    const teacherCandidate = $("a, span, div")
      .filter((_, el) => {
        const t = $(el).text().trim();
        return t.endsWith("老師") && t.length <= 20;
      })
      .first()
      .text()
      .trim();
    if (teacherCandidate) teacher = teacherCandidate;

    const pageText = $("body").text().replace(/\s+/g, " ");

    const profit = (() => {
      const idx = pageText.indexOf("淨利");
      if (idx < 0) return "";
      return firstMatchNumber(pageText.slice(idx, idx + 90));
    })();

    const winRate = (() => {
      const idx = pageText.indexOf("勝率");
      if (idx < 0) return "";
      return firstMatchNumber(pageText.slice(idx, idx + 90));
    })();

    const rr = (() => {
      const idx = pageText.indexOf("風報比");
      if (idx < 0) return "";
      return firstMatchNumber(pageText.slice(idx, idx + 90));
    })();

    // 標籤（容錯抓短字串）
    const tags = [];
    $("a, span, div").each((_, el) => {
      const t = $(el).text().trim();
      if (!t || t.length > 18) return;
      if (t === "查看詳情" || t === "策略市集") return;
      if (/^(#)?[A-Z]{2,6}\d?$/.test(t) || /台指|加權|櫃買|期貨|策略|程式交易|自動交易/.test(t)) {
        tags.push(t.replace(/^#/, ""));
      }
    });
    const tagText = [...new Set(tags)].slice(0, 12).join(" / ");

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

/* ------------------ resume ------------------ */

function loadExistingRows() {
  if (!fs.existsSync(OUT_XLSX)) return [];
  try {
    const wb = XLSX.readFile(OUT_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws);
    const cleaned = [];
    for (const r of rows) {
      const id = Number(r.ID);
      if (!Number.isFinite(id)) continue;
      cleaned.push({
        ID: id,
        名稱: String(r.名稱 ?? "").trim(),
        老師: String(r.老師 ?? "").trim(),
        淨利: String(r.淨利 ?? "").trim(),
        勝率: String(r.勝率 ?? "").trim(),
        風報比: String(r.風報比 ?? "").trim(),
        標籤: String(r.標籤 ?? "").trim(),
        網址: String(r.網址 ?? "").trim(),
      });
    }
    return cleaned;
  } catch {
    return [];
  }
}

/* ------------------ safe write (EBUSY-safe) ------------------ */

function writeCsv(rows, csvPath) {
  const headers = ["ID", "名稱", "老師", "淨利", "勝率", "風報比", "標籤", "網址"];
  const escapeCsv = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv =
    headers.join(",") +
    "\n" +
    rows.map((r) => headers.map((h) => escapeCsv(r[h])).join(",")).join("\n");
  fs.writeFileSync(csvPath, csv, "utf8");
}

function writeXlsx(rows, xlsxPath) {
  const headers = ["ID", "名稱", "老師", "淨利", "勝率", "風報比", "標籤", "網址"];
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  ws["!cols"] = [
    { wch: 8 }, { wch: 50 }, { wch: 18 }, { wch: 14 },
    { wch: 10 }, { wch: 10 }, { wch: 35 }, { wch: 70 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "products");
  XLSX.writeFile(wb, xlsxPath);
}

function safeWriteOutputs(rows) {
  // 去重 + 排序
  const uniq = new Map();
  for (const r of rows) uniq.set(Number(r.ID), r);
  const sorted = [...uniq.values()].sort((a, b) => Number(a.ID) - Number(b.ID));

  const stamp = nowStamp();
  const tmpXlsx = `product_total.tmp.${stamp}.xlsx`;
  const tmpCsv = `product_total.tmp.${stamp}.csv`;

  try {
    writeXlsx(sorted, tmpXlsx);
    writeCsv(sorted, tmpCsv);

    try {
      fs.renameSync(tmpXlsx, OUT_XLSX);
      fs.renameSync(tmpCsv, OUT_CSV);
      return { ok: true, locked: false, xlsx: OUT_XLSX, csv: OUT_CSV, count: sorted.length };
    } catch (e) {
      const fallbackXlsx = `product_total_${stamp}.xlsx`;
      const fallbackCsv = `product_total_${stamp}.csv`;
      fs.renameSync(tmpXlsx, fallbackXlsx);
      fs.renameSync(tmpCsv, fallbackCsv);
      return { ok: false, locked: true, xlsx: fallbackXlsx, csv: fallbackCsv, count: sorted.length, err: e?.code || "" };
    }
  } catch (e) {
    return { ok: false, locked: false, xlsx: "", csv: "", count: sorted.length, err: e?.code || e?.message || "" };
  }
}

/* ------------------ probe upper bound ------------------ */

async function blockHasAnySuccess(fromId, toId, tpl, doneSet) {
  if (toId < fromId) return false;

  const candidates = new Set([fromId, toId]);
  const step = 10;
  for (let i = fromId; i <= toId; i += step) candidates.add(i);

  const sorted = [...candidates].sort((a, b) => a - b);
  for (const id of sorted) {
    if (doneSet?.has(id)) return true;
    const r = await fetchProduct(id, tpl);
    await sleep(200);
    if (r) return true;
  }
  return false;
}

async function probeUpperBound(startId, tpl, doneSet) {
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
      return hi;
    }
  }
  return Math.min(lastSeen + PROBE_BLOCK, HARD_MAX_ID);
}

/* ------------------ main ------------------ */

function fmtRate(n) {
  if (!Number.isFinite(n)) return "";
  return n.toFixed(2);
}

async function main() {
  let input = process.argv[2];
  if (!input) input = await askUrl();
  input = toAbs(input);

  let startId, template;

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

  const existing = loadExistingRows();
  const rows = [...existing];
  const done = new Set(rows.map((r) => Number(r.ID)).filter((n) => Number.isFinite(n)));

  console.log(`✔ 已有舊檔：${rows.length} 筆（將自動續抓、跳過已完成 ID）`);

  console.log("🔍 探測合理上界中（加速用）…");
  const upperGuess = await probeUpperBound(startId, template, done);
  console.log(`✔ 探測上界候選：ID ~ ${upperGuess}`);

  console.log("🚀 開始正式掃描（不中斷、不怕斷號）…");
  console.log("ℹ 執行中按 Enter：立刻顯示進度並強制回寫 product_total.xlsx\n");

  // ===== 心跳/進度統計 =====
  const startedAt = Date.now();
  let lastBeat = Date.now();

  let lastSuccessId = 0;
  let addedSinceSave = 0;

  let missStreak = 0;
  let checked = 0; // 掃過多少 ID（含已跳過）
  let fetched = 0; // 真正發 request 的次數（不含 done skip）
  let currentId = startId;

  // ===== 手動按 Enter 強制存檔（避免重入） =====
  let manualSaving = false;
  let manualPending = false;

  function statusLine(prefix = "…掃描中") {
    return `${prefix}：ID=${currentId}（連續空號 ${missStreak}）最後成功=${lastSuccessId || "(尚未)"} 目前筆數=${rows.length}`;
  }

  function doSave(tag = "手動存檔") {
    const ret = safeWriteOutputs(rows);
    if (ret.ok) {
      console.log(`💾 ${tag}完成：${ret.xlsx}（${ret.count} 筆）`);
    } else if (ret.locked) {
      console.log(`⚠️ ${tag}：主檔被占用（可能 Excel 開著），改存：${ret.xlsx}（${ret.count} 筆）`);
    } else {
      console.log(`⚠️ ${tag}失敗但不中斷：${ret.err}`);
    }
  }

  // 建立一個 readline 監聽 Enter
  const rlKey = readline.createInterface({ input: process.stdin, output: process.stdout });
  rlKey.on("line", () => {
    // 你要的：按 Enter 出現那行，然後立刻回寫
    console.log(statusLine("…掃描中"));

    if (manualSaving) {
      manualPending = true;
      return;
    }

    manualSaving = true;
    try {
      doSave("手動回寫");
    } finally {
      manualSaving = false;
      if (manualPending) {
        manualPending = false;
        // 若連按很多次 Enter，只再補存一次
        manualSaving = true;
        try {
          doSave("手動回寫(補存)");
        } finally {
          manualSaving = false;
        }
      }
    }
  });

  try {
    for (let id = startId; id <= HARD_MAX_ID; id++) {
      currentId = id;
      checked++;

      // 已完成直接跳過，但仍更新 lastSuccessId（代表成功存在）
      if (done.has(id)) {
        lastSuccessId = Math.max(lastSuccessId, id);

        // 心跳：即使都在 skip 也要印
        if (Date.now() - lastBeat >= HEARTBEAT_SEC * 1000) {
          const elapsedSec = (Date.now() - startedAt) / 1000;
          const speed = fetched > 0 ? fetched / elapsedSec : 0;
          console.log(
            `⏳ still running... 目前掃到 ID=${id}（skip）最後成功=${lastSuccessId || "(尚未)"} 目前筆數=${rows.length} 速度=${fmtRate(speed)} req/s 連續空號=${missStreak}`
          );
          lastBeat = Date.now();
        }
        continue;
      }

      fetched++;
      const r = await fetchProduct(id, template);

      if (r) {
        rows.push(r);
        done.add(id);
        lastSuccessId = id;
        addedSinceSave++;
        missStreak = 0;
        console.log(`✔ ID=${id} ${r.名稱}`);
      } else {
        missStreak++;
        if (missStreak % SHOW_MISS_EVERY === 0) {
          console.log(`…掃描中：ID=${id}（連續空號 ${missStreak}）最後成功=${lastSuccessId || "(尚未)"} 目前筆數=${rows.length}`);
        }
      }

      // ✅ 結尾判斷：先跑到 upperGuess；之後超過 GAP_LIMIT 都沒成功就停
      if (id >= upperGuess && lastSuccessId > 0 && id - lastSuccessId >= GAP_LIMIT) {
        console.log(`\n🛑 超過最後成功 ID=${lastSuccessId} 後 ${GAP_LIMIT} 筆皆無資料，停止`);
        break;
      }

      // ✅ 心跳：固定顯示目前掃到哪（不管有沒有成功）
      if (Date.now() - lastBeat >= HEARTBEAT_SEC * 1000) {
        const elapsedSec = (Date.now() - startedAt) / 1000;
        const speed = fetched > 0 ? fetched / elapsedSec : 0;
        console.log(
          `⏳ still running... 目前掃到 ID=${id}，最後成功=${lastSuccessId || "(尚未)"} 目前筆數=${rows.length} 速度=${fmtRate(speed)} req/s 連續空號=${missStreak}`
        );
        lastBeat = Date.now();
      }

      // ✅ 自動存檔：就算 Excel 開著也不中斷
      if (addedSinceSave >= SAVE_EVERY) {
        doSave("自動存檔");
        addedSinceSave = 0;
      }

      await sleep(SLEEP_MS);
    }
  } finally {
    // 關掉 Enter 監聽，避免程式結束還卡著 stdin
    rlKey.close();
  }

  // 最後再存一次
  doSave("最後存檔");

  console.log("\n✅ 結束");
}

main().catch((e) => {
  console.error("❌", e?.message || e);
  process.exit(1);
});
