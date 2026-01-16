// 檔名：P爬商品總表_商品MC.mjs
import axios from "axios";
import * as cheerio from "cheerio";
import fs from "fs";
import readline from "readline";
import XLSX from "xlsx";

const HEADERS = { "User-Agent": "Mozilla/5.0" };

const SLEEP_MS = 450;
const SAVE_EVERY = 10;

const HARD_MAX_ID = 20000;   // 保險上限
const GAP_LIMIT = 2000;      // ✅ 放大：避免 193 後面跳很遠導致漏抓

const PROBE_BLOCK = 200;
const PROBE_STEP_START = 200;
const PROBE_MAX_ROUNDS = 20;

const OUT_XLSX = "product_total.xlsx";
const OUT_CSV = "product_total.csv";

/* ------------------ utils ------------------ */

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function askUrl() {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question("請輸入第一個商品網址或根網址（例如 https://mcsm.multicharts.com.tw/ 或 .../product/info/1）：\n", ans => {
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
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
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

    // 標籤：簡單抓短字串（容錯）
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
  const csv = headers.join(",") + "\n" + rows.map(r => headers.map(h => escapeCsv(r[h])).join(",")).join("\n");
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
  // 先去重+排序
  const uniq = new Map();
  for (const r of rows) uniq.set(Number(r.ID), r);
  const sorted = [...uniq.values()].sort((a, b) => Number(a.ID) - Number(b.ID));

  // 原子寫入：先寫 tmp，再 rename
  const stamp = nowStamp();
  const tmpXlsx = `product_total.tmp.${stamp}.xlsx`;
  const tmpCsv  = `product_total.tmp.${stamp}.csv`;

  try {
    writeXlsx(sorted, tmpXlsx);
    writeCsv(sorted, tmpCsv);

    // 先嘗試覆蓋正式檔
    try {
      fs.renameSync(tmpXlsx, OUT_XLSX);
      fs.renameSync(tmpCsv, OUT_CSV);
      return { ok: true, xlsx: OUT_XLSX, csv: OUT_CSV, count: sorted.length };
    } catch (e) {
      // 若正式檔被鎖住（EBUSY/EPERM），改寫成 timestamp 檔名
      const fallbackXlsx = `product_total_${stamp}.xlsx`;
      const fallbackCsv  = `product_total_${stamp}.csv`;

      fs.renameSync(tmpXlsx, fallbackXlsx);
      fs.renameSync(tmpCsv, fallbackCsv);

      return { ok: false, locked: true, xlsx: fallbackXlsx, csv: fallbackCsv, count: sorted.length, err: e?.code || "" };
    }
  } catch (e) {
    // 如果 tmp 寫入就失敗，也不要中斷主程式
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
  const done = new Set(rows.map(r => Number(r.ID)).filter(n => Number.isFinite(n)));
  console.log(`✔ 已有舊檔：${rows.length} 筆（將自動續抓、跳過已完成 ID）`);

  console.log("🔍 探測合理上界中（加速用）…");
  const upperGuess = await probeUpperBound(startId, template, done);
  console.log(`✔ 探測上界候選：ID ~ ${upperGuess}`);

  console.log("🚀 開始正式掃描（不中斷、不怕斷號）…\n");

  let lastSuccessId = 0;
  let addedSinceSave = 0;

  for (let id = startId; id <= HARD_MAX_ID; id++) {
    if (done.has(id)) {
      lastSuccessId = Math.max(lastSuccessId, id);
      continue;
    }

    const r = await fetchProduct(id, template);
    if (r) {
      rows.push(r);
      done.add(id);
      lastSuccessId = id;
      addedSinceSave++;

      console.log(`✔ ID=${id} ${r.名稱}`);
    }

    // ✅ 結尾：先跑到 upperGuess；之後用 GAP_LIMIT 判斷真的結束
    if (id >= upperGuess && lastSuccessId > 0 && id - lastSuccessId >= GAP_LIMIT) {
      console.log(`\n🛑 超過最後成功 ID=${lastSuccessId} 後 ${GAP_LIMIT} 筆皆無資料，停止`);
      break;
    }

    // ✅ 自動存檔：就算 Excel 開著也不中斷
    if (addedSinceSave >= SAVE_EVERY) {
      const ret = safeWriteOutputs(rows);
      if (ret.ok) {
        console.log(`💾 已存檔：${ret.xlsx}（${ret.count} 筆）`);
      } else if (ret.locked) {
        console.log(`⚠️ 主要檔案被占用（可能 Excel 開著），改存：${ret.xlsx}（${ret.count} 筆）`);
      } else {
        console.log(`⚠️ 存檔失敗但繼續爬：${ret.err}`);
      }
      addedSinceSave = 0;
    }

    await sleep(SLEEP_MS);
  }

  // 最後再存一次
  const final = safeWriteOutputs(rows);
  if (final.ok) {
    console.log(`\n✅ 完成，共 ${final.count} 筆`);
    console.log(`📄 輸出：${final.xlsx}`);
    console.log(`📄 輸出：${final.csv}`);
  } else if (final.locked) {
    console.log(`\n✅ 完成，但主檔案被占用（可能 Excel 開著）`);
    console.log(`📄 改存：${final.xlsx}（${final.count} 筆）`);
    console.log(`📄 改存：${final.csv}`);
  } else {
    console.log(`\n✅ 完成，但最後存檔失敗：${final.err}`);
    console.log("（你可以關掉 Excel 後再重跑一次，它會續抓/續存）");
  }
}

main().catch(e => {
  console.error("❌", e?.message || e);
  process.exit(1);
});
