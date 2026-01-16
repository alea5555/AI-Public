import axios from "axios";
import * as cheerio from "cheerio";
import readline from "readline";
import * as XLSX from "xlsx";

const SLEEP_MS = 600;
const MAX_MISS = 10;
const HEADERS = { "User-Agent": "Mozilla/5.0" };

/* ---------- CLI ---------- */

function askFirstUrl() {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question("請輸入第一個商品網址（例如 .../product/info/1）：\n", ans => {
      rl.close();
      resolve((ans || "").trim());
    });
  });
}

/* ---------- 解析網址 ---------- */

function parseUrlTemplate(firstUrl) {
  let u;
  try {
    u = new URL(firstUrl);
  } catch {
    throw new Error("網址格式錯誤");
  }

  const parts = u.pathname.split("/").filter(Boolean);
  const last = parts.at(-1);

  if (!/^\d+$/.test(last)) {
    throw new Error("網址最後一段必須是數字 ID");
  }

  const startId = parseInt(last, 10);
  const basePath = parts.slice(0, -1).join("/");

  const template = `${u.origin}/${basePath}/{id}`;

  return { template, startId };
}

/* ---------- 抓單一商品 ---------- */

async function fetchOne(id, urlTpl) {
  const url = urlTpl.replace("{id}", id);

  try {
    const res = await axios.get(url, { headers: HEADERS, timeout: 15000 });
    const $ = cheerio.load(res.data);

    const title =
      $("h1").first().text().trim() ||
      $("h2").first().text().trim();

    if (!title) return null;

    return {
      ID: id,
      名稱: title,
      網址: url
    };
  } catch {
    return null;
  }
}

/* ---------- Excel ---------- */

function writeXlsx(rows) {
  const ws = XLSX.utils.json_to_sheet(rows, {
    header: ["ID", "名稱", "網址"]
  });
  ws["!cols"] = [{ wch: 8 }, { wch: 50 }, { wch: 70 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "products");
  XLSX.writeFile(wb, "product_total.xlsx");
}

/* ---------- 主流程 ---------- */

async function main() {
  let firstUrl = process.argv[2];
  if (!firstUrl) firstUrl = await askFirstUrl();

  const { template, startId } = parseUrlTemplate(firstUrl);

  console.log(`✔ 偵測連號網址：${template}`);
  console.log(`✔ 起始 ID：${startId}\n`);

  const rows = [];
  let miss = 0;
  let id = startId;

  while (true) {
    const r = await fetchOne(id, template);

    if (r) {
      rows.push(r);
      miss = 0;
      console.log(`✔ ID=${id} ${r.名稱}`);
    } else {
      miss++;
      console.log(`✖ ID=${id} 無資料 (${miss}/${MAX_MISS})`);
    }

    if (miss >= MAX_MISS) break;

    id++;
    await new Promise(r => setTimeout(r, SLEEP_MS));
  }

  writeXlsx(rows);

  console.log(`\n✅ 完成`);
  console.log(`📦 共 ${rows.length} 筆`);
  console.log(`📄 輸出：product_total.xlsx`);
}

main().catch(e => {
  console.error("❌", e.message);
});
